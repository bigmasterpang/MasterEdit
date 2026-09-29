import { LazyStore } from "@tauri-apps/plugin-store";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow, LogicalSize } from "@tauri-apps/api/window";
import { createDoc, docFromPayload, useAppStore } from "../stores/appStore";
import { pickSettings, useSettingsStore } from "../stores/settingsStore";
import type {
  BinaryPayload,
  FilePayload,
  PdfHighlight,
  PdfNote,
  Settings,
  SheetEdit,
  SpreadsheetInfo,
  ViewMode,
} from "../types";
import { LARGE_FILE_BYTES } from "./constants";
import { isPdfPath, isSpreadsheetPath, normalizeSlashes, textReadOnlyLimit } from "./filePath";
import { debounce } from "./timing";

/** 是否运行在 Tauri 环境中（纯浏览器打开 vite 页面时跳过持久化） */
export const isTauri =
  typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;

interface UiState {
  viewMode: ViewMode;
  outlineVisible: boolean;
  windowWidth: number;
  windowHeight: number;
  split?: boolean;
  splitRatio?: number;
  activePane?: 0 | 1;
}

/**
 * 插件存储文件名（位于 `%APPDATA%\com.masterpang.mastermd`）。
 * 改名后由 Rust 端 `commands/legacy.rs` 在启动时把旧的 `mastermd-store.json`
 * 一次性复制过来，因此这里可以直接使用新名字而不会让老用户丢设置与 PDF 批注。
 */
const STORE_FILE = "masteredit-store.json";
let file: LazyStore | null = null;
let uiCache: UiState = {
  viewMode: "split",
  outlineVisible: true,
  windowWidth: 0,
  windowHeight: 0,
};

async function readWindowSize(): Promise<{ width: number; height: number }> {
  try {
    const win = getCurrentWindow();
    if (await win.isMaximized()) return { width: 0, height: 0 };
    const size = await win.innerSize();
    const scale = await win.scaleFactor();
    return {
      width: Math.round(size.width / scale),
      height: Math.round(size.height / scale),
    };
  } catch {
    return { width: 0, height: 0 };
  }
}

const persistUi = debounce(() => {
  void file?.set("ui", uiCache);
}, 400);

async function captureUi(): Promise<void> {
  const app = useAppStore.getState();
  const size = await readWindowSize();
  uiCache = {
    viewMode: app.viewMode,
    outlineVisible: app.outlineVisible,
    windowWidth: size.width || uiCache.windowWidth,
    windowHeight: size.height || uiCache.windowHeight,
    split: app.layout.split,
    splitRatio: app.layout.ratio,
    activePane: app.layout.activePane,
  };
  persistUi();
}

const persistSettings = debounce(() => {
  void file?.set("settings", pickSettings(useSettingsStore.getState()));
}, 400);

/** 应用启动时初始化持久化：恢复设置 / 界面状态，并订阅后续变化 */
export async function initPersistence(): Promise<void> {
  if (!isTauri) {
    useSettingsStore.setState({ loaded: true });
    return;
  }
  try {
    file = new LazyStore(STORE_FILE);
    await file.init();

    const settings = await file.get<Partial<Settings>>("settings");
    if (settings) useSettingsStore.getState().hydrate(settings);
    else useSettingsStore.setState({ loaded: true });

    const ui = await file.get<Partial<UiState>>("ui");
    if (ui) {
      uiCache = { ...uiCache, ...ui };
      useAppStore.setState({
        viewMode: ui.viewMode ?? "split",
        outlineVisible: ui.outlineVisible ?? true,
        layout: {
          split: ui.split ?? false,
          activePane: ui.activePane ?? 0,
          ratio: typeof ui.splitRatio === "number" ? ui.splitRatio : 0.5,
        },
      });
      if (ui.windowWidth && ui.windowHeight) {
        try {
          await getCurrentWindow().setSize(
            new LogicalSize(ui.windowWidth, ui.windowHeight),
          );
        } catch {
          /* 尺寸恢复失败不影响使用 */
        }
      }
    }
  } catch (error) {
    console.error("初始化本地存储失败", error);
    useSettingsStore.setState({ loaded: true });
  }

  useSettingsStore.subscribe(persistSettings);
  useAppStore.subscribe((state, prev) => {
    if (
      state.viewMode !== prev.viewMode ||
      state.outlineVisible !== prev.outlineVisible ||
      state.layout.split !== prev.layout.split ||
      state.layout.ratio !== prev.layout.ratio ||
      state.layout.activePane !== prev.layout.activePane
    ) {
      void captureUi();
    }
  });
  window.addEventListener("resize", () => void captureUi());
}

/** 应用退出前保存一次窗口尺寸 */
export async function flushUiState(): Promise<void> {
  if (!isTauri || !file) return;
  await captureUi();
  // 会话是 1.2s 防抖写入的：关闭窗口时立即补写一次，
  // 否则「刚打开的标签页」可能因为防抖还没触发而丢失
  await writeSession();
  try {
    await file.save();
  } catch {
    /* ignore */
  }
}

/* ------------------------------------------------------------------ */
/* PDF 标注数据持久化：高亮、批注、便签、底色、阅读位置                */
/* ------------------------------------------------------------------ */

export interface StoredPdfAnnotations {
  highlights: PdfHighlight[];
  notes: PdfNote[];
  paperTheme?: string;
  currentPage?: number;
}

export async function loadPdfAnnotations(filePath: string): Promise<StoredPdfAnnotations | null> {
  if (!isTauri) return null;
  try {
    if (!file) {
      file = new LazyStore(STORE_FILE);
      await file.init();
    }
    const key = `pdf_anno:${normalizeSlashes(filePath).toLowerCase()}`;
    const data = await file.get<StoredPdfAnnotations>(key);
    return data ?? null;
  } catch (error) {
    console.error("读取 PDF 标注失败", error);
    return null;
  }
}

export async function savePdfAnnotations(
  filePath: string,
  annotations: StoredPdfAnnotations,
): Promise<void> {
  if (!isTauri) return;
  try {
    if (!file) {
      file = new LazyStore(STORE_FILE);
      await file.init();
    }
    const key = `pdf_anno:${normalizeSlashes(filePath).toLowerCase()}`;
    await file.set(key, annotations);
    await file.save();
  } catch (error) {
    console.error("保存 PDF 标注失败", error);
  }
}

/* ------------------------------------------------------------------ */
/* 会话恢复：避免意外重载（如误按刷新）导致未保存内容丢失              */
/* ------------------------------------------------------------------ */

interface SessionPayload {
  /** 已保存文档的路径及窗格（用于恢复标签页，兼容旧版 string 数组） */
  paths: Array<string | { path: string; pane?: 0 | 1 }>;
  /** 未保存 / 未命名文档的内容及窗格 */
  unsaved: Array<{
    path: string | null;
    content: string;
    pane?: 0 | 1;
    docType?: string;
    /** 表格未提交的单元格编辑（content 为空时靠它恢复） */
    sheetEdits?: SheetEdit[];
  }>;
}

const SESSION_TAB_LIMIT = 8;
const SESSION_CONTENT_LIMIT = 512 * 1024;
/** 会话恢复时单个文件的大小上限：超过则连标签一起放弃，避免启动时卡死 */
const SESSION_RESTORE_MAX_BYTES = 32 * 1024 * 1024;

const persistSession = debounce(() => {
  void writeSession();
}, 1200);

async function writeSession(): Promise<void> {
  if (!file) return;
  try {
    const { docs } = useAppStore.getState();
    /** 内容过大（或已是 PDF）的文档不能靠「按路径恢复」找回未保存内容 */
    const oversized = (doc: (typeof docs)[number]) => doc.content.length > SESSION_CONTENT_LIMIT;

    const unsaved = docs
      .filter((doc) => doc.isDirty || !doc.filePath)
      .filter((doc) => doc.docType !== "pdf" && !isPdfPath(doc.filePath))
      .filter((doc) => !oversized(doc))
      .slice(0, SESSION_TAB_LIMIT)
      .map((doc) => ({
        path: doc.filePath,
        content: doc.content,
        pane: doc.pane ?? 0,
        // 不记 docType 的话，未命名的纯文本文档恢复后会变成 Markdown
        docType: doc.docType,
        // 表格的未提交编辑（体积很小，不受 content 上限影响）
        sheetEdits: doc.sheetEdits,
      }));

    const unsavedPaths = new Set(
      unsaved
        .map((entry) => entry.path?.toLowerCase())
        .filter((path): path is string => Boolean(path)),
    );

    const payload: SessionPayload = {
      paths: docs
        .filter((doc) => doc.filePath)
        // 只有「超大 **且** 有未保存修改」的文档不能按路径恢复：从磁盘重读会静默覆盖编辑。
        // 干净的超大文档（例如大 CSV）按路径恢复是安全的 —— 否则每次重启都会丢标签。
        .filter((doc) => !(oversized(doc) && doc.isDirty))
        // 同上：有未保存内容但被 unsaved 截断掉的文档也不能走路径恢复
        .filter(
          (doc) =>
            !(
              doc.isDirty &&
              doc.docType !== "pdf" &&
              !isPdfPath(doc.filePath) &&
              !unsavedPaths.has((doc.filePath as string).toLowerCase())
            ),
        )
        .map((doc) => ({ path: doc.filePath as string, pane: doc.pane ?? 0 }))
        .slice(0, SESSION_TAB_LIMIT),
      unsaved,
    };
    await file.set("session", payload);
  } catch (error) {
    console.error("保存会话失败", error);
  }
}

/** 启动时恢复上次未保存的内容与打开的标签页 */
export async function restoreSession(): Promise<void> {
  if (!isTauri || !file) return;
  try {
    const payload = await file.get<SessionPayload>("session");
    if (!payload) return;
    const state = useAppStore.getState();
    const opened = new Set(
      state.docs
        .map((doc) => doc.filePath?.toLowerCase())
        .filter((path): path is string => Boolean(path)),
    );

    // 1. 未保存 / 未命名文档优先恢复（过滤掉 PDF）
    for (const entry of payload.unsaved ?? []) {
      // 表格的未保存内容在 sheetEdits 里（content 为空），不能按「内容为空」跳过
      const sheetEdits = entry.sheetEdits ?? [];
      if (!entry.content && sheetEdits.length === 0) continue;
      if (entry.path && isPdfPath(entry.path)) continue;

      // 表格：按路径重新解析工作簿，再把未提交的单元格编辑挂回去
      if (sheetEdits.length > 0 && entry.path && isSpreadsheetPath(entry.path)) {
        try {
          const info = await invoke<SpreadsheetInfo>("spreadsheet_info", { path: entry.path });
          state.addDoc(
            createDoc({
              filePath: info.path,
              docType: "spreadsheet",
              content: "",
              savedContent: "",
              isDirty: true,
              readOnly: !info.editable || info.size > LARGE_FILE_BYTES,
              encrypted: info.encrypted,
              sheetEdits,
              pane: entry.pane ?? 0,
              modifiedAt: info.modifiedAt,
              size: info.size,
            }),
          );
          opened.add(entry.path.toLowerCase());
          void invoke("watch_file", { path: entry.path }).catch(() => undefined);
        } catch {
          /* 打不开的表格跳过 */
        }
        continue;
      }

      const doc = createDoc({
        filePath: entry.path,
        content: entry.content,
        savedContent: entry.path ? "" : entry.content,
        isDirty: Boolean(entry.path),
        pane: entry.pane ?? 0,
        // 旧会话没有 docType：有路径时按路径推断，没路径的旧数据按 Markdown 处理
        docType:
          entry.docType === "blank" || entry.docType === "pdf"
            ? entry.docType
            : entry.path
              ? undefined
              : "markdown",
      });
      if (entry.path) {
        try {
          const disk = await invoke<FilePayload>("read_markdown_file", { path: entry.path });
          doc.savedContent = disk.content;
          doc.modifiedAt = disk.modifiedAt;
          doc.size = disk.size;
          doc.isDirty = disk.content !== entry.content;
          opened.add(entry.path.toLowerCase());
          // 恢复的标签页同样要监听外部变更，否则下一次保存会覆盖别人的改动
          void invoke("watch_file", { path: entry.path }).catch(() => undefined);
        } catch {
          // 文件已不存在：保留内容，等待用户另存为
          doc.isDirty = true;
        }
      }
      state.addDoc(doc);
    }

    // 2. 其余已保存标签页（支持恢复 PDF 及标注）
    for (const item of payload.paths ?? []) {
      const path = typeof item === "string" ? item : item.path;
      const pane = typeof item === "string" ? 0 : item.pane ?? 0;
      if (opened.has(path.toLowerCase())) continue;
      try {
        // 电子表格（只读）：按路径恢复，内容由 Rust 侧按需解析
        if (isSpreadsheetPath(path)) {
          const info = await invoke<SpreadsheetInfo>("spreadsheet_info", { path });
          state.addDoc(
            createDoc({
              filePath: info.path,
              docType: "spreadsheet",
              pane,
              content: "",
              savedContent: "",
              isDirty: false,
              // 与手动打开一致：只有 .xlsx 且体积不大时可编辑
              readOnly: !info.editable || info.size > LARGE_FILE_BYTES,
              encrypted: info.encrypted,
              modifiedAt: info.modifiedAt,
              size: info.size,
            }),
          );
          opened.add(path.toLowerCase());
          void invoke("watch_file", { path }).catch(() => undefined);
          continue;
        }

        if (isPdfPath(path)) {
          const payloadData = await invoke<BinaryPayload>("read_binary_file", { path });
          const anno = await loadPdfAnnotations(path);
          const restoredDoc = createDoc({
            filePath: payloadData.path,
            docType: "pdf",
            pane,
            pdfBase64: payloadData.dataBase64,
            savedPdfBase64: payloadData.dataBase64,
            cleanPdfBase64: payloadData.dataBase64,
            pdfHighlights: anno?.highlights ?? [],
            pdfNotes: anno?.notes ?? [],
            pdfPaperTheme: anno?.paperTheme ?? "white",
            pdfCurrentPage: anno?.currentPage ?? 1,
            modifiedAt: payloadData.modifiedAt,
            size: payloadData.size,
            encrypted: payloadData.encrypted ?? false,
            encryptedHeader: payloadData.encryptedHeader ?? null,
            readOnly: false,
          });
          state.addDoc(restoredDoc);
          opened.add(path.toLowerCase());
          // PDF 标签页同样需要监听（批注侧车 + 外部改动检测）
          void invoke("watch_file", { path }).catch(() => undefined);
          continue;
        }

        const payloadData = await invoke<FilePayload>("read_markdown_file", { path });
        // 超大文件（如大 CSV / 日志）按只读恢复，而不是直接丢掉标签：
        // 与手动打开大文件的行为一致，避免启动时误编辑超大文件
        if (payloadData.size > SESSION_RESTORE_MAX_BYTES) continue;
        const restoredDoc = docFromPayload(payloadData);
        restoredDoc.pane = pane;
        if (payloadData.size > textReadOnlyLimit(path)) restoredDoc.readOnly = true;
        state.addDoc(restoredDoc);
        opened.add(path.toLowerCase());
        // 会话恢复出来的标签页必须重新注册文件监听：openPath 只在「打开」时注册，
        // 漏掉这里会导致外部修改无提示、下次保存直接覆盖
        void invoke("watch_file", { path }).catch(() => undefined);
      } catch {
        /* 打不开的文件跳过 */
      }
    }
  } catch (error) {
    console.error("恢复会话失败", error);
  }
}

/** 开始监听文档变化并持久化会话 */
export function startSessionTracking(): void {
  useAppStore.subscribe((state, prev) => {
    if (state.docs !== prev.docs) persistSession();
  });
}
