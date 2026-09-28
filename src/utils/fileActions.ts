import { invoke } from "@tauri-apps/api/core";
import { open as openDialog } from "@tauri-apps/plugin-dialog";
import { openUrl } from "@tauri-apps/plugin-opener";
import { createDoc, docFromPayload, getActiveDoc, getDocById, useAppStore } from "../stores/appStore";
import { askConfirm, askUnsaved, showMessage } from "../stores/dialogStore";
import type { BinaryPayload, FilePayload } from "../types";
import { loadPdfAnnotations, savePdfAnnotations } from "./persist";
import {
  EMPTY_DOC_PLACEHOLDER,
  LARGE_FILE_BYTES,
  OPEN_DIALOG_FILTERS,
} from "./constants";
import { fileName, isMarkdownPath, isOpenablePath, isPdfPath, samePath } from "./filePath";
import { formatBytes } from "./timing";

/** 展示用名称 */
export function displayName(doc: { filePath: string | null }): string {
  return doc.filePath ? fileName(doc.filePath) : "未命名文档";
}

function utf8Size(text: string): number {
  return new TextEncoder().encode(text).length;
}

/* ------------------------------ 自身写入标记 -------------------------- */

let lastSelfWriteAt = 0;

/** 记录一次由本应用发起的写入，用于忽略监听器回传的自身事件 */
export function markSelfWrite(): void {
  lastSelfWriteAt = Date.now();
}

export function isRecentSelfWrite(windowMs = 1200): boolean {
  return Date.now() - lastSelfWriteAt < windowMs;
}

/* ------------------------------ 最近文件 ------------------------------ */

export async function loadRecentFiles(): Promise<void> {
  try {
    const list = await invoke<string[]>("get_recent_files");
    useAppStore.getState().setRecentFiles(list);
  } catch (error) {
    console.error("读取最近文件失败", error);
  }
}

export async function addRecentFile(path: string): Promise<void> {
  try {
    const list = await invoke<string[]>("add_recent_file", { path });
    useAppStore.getState().setRecentFiles(list);
  } catch (error) {
    console.error("记录最近文件失败", error);
  }
}

export async function removeRecentFile(path: string): Promise<void> {
  try {
    const list = await invoke<string[]>("remove_recent_file", { path });
    useAppStore.getState().setRecentFiles(list);
  } catch (error) {
    console.error("移除最近文件失败", error);
  }
}

export async function clearRecentFiles(): Promise<void> {
  try {
    const list = await invoke<string[]>("clear_recent_files");
    useAppStore.getState().setRecentFiles(list);
  } catch (error) {
    console.error("清空最近文件失败", error);
  }
}

/* ------------------------------ 文件监听 ------------------------------ */

export async function watchFile(path: string): Promise<void> {
  try {
    await invoke("watch_file", { path });
  } catch (error) {
    console.error("启动文件监听失败", error);
  }
}

export async function unwatchFile(path: string): Promise<void> {
  try {
    await invoke("unwatch_file", { path });
  } catch (error) {
    console.error("停止文件监听失败", error);
  }
}

/* ------------------------------ 打开文件 ------------------------------ */

/** 处理未保存变更；返回 false 表示用户取消 */
async function ensureNoDirty(doc: { id: string; filePath: string | null; isDirty: boolean } | null): Promise<boolean> {
  if (!doc || !doc.isDirty) return true;
  const choice = await askUnsaved(displayName(doc));
  if (choice === "cancel") return false;
  if (choice === "save") return await saveDoc(doc.id);
  return true;
}

export async function openPath(path: string, targetPane?: 0 | 1): Promise<boolean> {
  const currentActivePane = useAppStore.getState().layout.activePane ?? 0;
  const effectivePane = targetPane !== undefined ? targetPane : currentActivePane;

  if (!(await ensureNoDirty(getActiveDoc()))) return false;

  const existing = useAppStore
    .getState()
    .docs.find((d) => samePath(d.filePath, path));
  if (existing) {
    useAppStore.getState().activateDoc(existing.id, effectivePane);
    return true;
  }

  try {
    if (isPdfPath(path)) {
      const payload = await invoke<BinaryPayload>("read_binary_file", { path });
      const anno = await loadPdfAnnotations(path);
      const doc = createDoc({
        filePath: payload.path,
        docType: "pdf",
        pane: effectivePane,
        pdfBase64: payload.dataBase64,
        savedPdfBase64: payload.dataBase64,
        cleanPdfBase64: payload.dataBase64,
        pdfHighlights: anno?.highlights ?? [],
        pdfNotes: anno?.notes ?? [],
        pdfPaperTheme: anno?.paperTheme ?? "white",
        pdfCurrentPage: anno?.currentPage ?? 1,
        modifiedAt: payload.modifiedAt,
        size: payload.size,
        encrypted: payload.encrypted ?? false,
        encryptedHeader: payload.encryptedHeader ?? null,
        readOnly: false,
      });
      useAppStore.getState().addDoc(doc);
      void addRecentFile(payload.path);
      void watchFile(payload.path);
      return true;
    }

    const payload = await invoke<FilePayload>("read_markdown_file", { path });
    if (!payload.encrypted && payload.size > LARGE_FILE_BYTES) {
      const ok = await askConfirm({
        title: "文件较大",
        message: `「${fileName(payload.path)}」大小为 ${formatBytes(
          payload.size,
        )}，将以只读方式打开以保证流畅度。`,
        confirmText: "只读打开",
      });
      if (!ok) return false;
    }
    const doc = docFromPayload(payload, effectivePane);
    useAppStore.getState().addDoc(doc);
    // 非 Markdown 文件（代码/纯文本）固定以源码模式打开
    if (!isMarkdownPath(payload.path)) {
      useAppStore.getState().setViewMode("source");
    }
    void addRecentFile(payload.path);
    void watchFile(payload.path);
    // 加密文档不弹窗提示：状态栏已有「已解密」标记，保存时自动按原格式加密写回
    return true;
  } catch (error) {
    await showMessage("打开失败", `无法打开文件：\n${path}\n\n${String(error)}`);
    void removeRecentFile(path);
    return false;
  }
}

export async function openFileDialog(targetPane?: 0 | 1): Promise<void> {
  try {
    const selected = await openDialog({
      multiple: false,
      directory: false,
      title: "打开文件",
      filters: OPEN_DIALOG_FILTERS,
    });
    if (typeof selected === "string") {
      await openPath(selected, targetPane);
    }
  } catch (error) {
    await showMessage("打开失败", String(error));
  }
}

/** 拖放打开：只打开第一个文件 */
export async function openDroppedPaths(paths: string[]): Promise<void> {
  if (paths.length === 0) return;
  const openable = paths.filter(isOpenablePath);
  if (openable.length === 0) {
    await showMessage(
      "不支持的文件类型",
      `目前支持 Markdown（.md/.markdown/.mdown）、纯文本（.txt/.log 等）与常见代码/配置文件（.json/.js/.ts/.py/.java/.sql/.yml 等）。\n\n拖入的文件：\n${paths
        .slice(0, 5)
        .join("\n")}`,
    );
    return;
  }
  if (openable.length > 1) {
    await showMessage(
      "暂不支持多开",
      `一次只能打开一个文件，将打开第一个：\n${fileName(openable[0])}`,
    );
  }
  await openPath(openable[0]);
}

/* ------------------------------ 编码与换行符 ------------------------------ */

/**
 * 切换文件编码（重新解释）：
 * 从磁盘按新编码重新读取，内容变化时保持「未保存」，保存后即以新编码写回。
 */
export async function setDocEncoding(id: string, encoding: string): Promise<void> {
  const doc = getDocById(id);
  if (!doc) return;
  if (!doc.filePath) {
    useAppStore.getState().patchDoc(id, { encoding });
    return;
  }
  if (doc.isDirty) {
    const ok = await askConfirm({
      title: "切换编码",
      message: "切换编码会重新读取磁盘文件，未保存的修改将丢失。是否继续？",
      confirmText: "继续切换",
    });
    if (!ok) return;
  }
  try {
    const payload = await invoke<FilePayload>("read_markdown_file", {
      path: doc.filePath,
      encoding,
    });
    // content 变化时 patchDoc 会自动重算 isDirty：不同即视为待保存
    useAppStore.getState().patchDoc(id, {
      content: payload.content,
      encoding,
      eol: payload.eol ?? doc.eol,
      size: payload.size,
      modifiedAt: payload.modifiedAt,
    });
  } catch (error) {
    await showMessage("切换编码失败", String(error));
  }
}

/** 切换换行符（保存时统一转换） */
export function setDocEol(id: string, eol: string): void {
  const doc = getDocById(id);
  if (!doc) return;
  const lf = doc.content.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const converted =
    eol === "crlf" ? lf.replace(/\n/g, "\r\n") : eol === "cr" ? lf.replace(/\n/g, "\r") : lf;
  useAppStore.getState().patchDoc(id, { eol, isDirty: converted !== doc.savedContent });
}

/* ------------------------------ 保存文件 ------------------------------ */

export async function saveDoc(id: string): Promise<boolean> {
  const doc = getDocById(id);
  if (!doc) return false;
  if (!doc.filePath) return saveDocAs(id);
  if (doc.readOnly) {
    await showMessage(
      "只读文档",
      "当前文档以只读方式打开（文件过大），请使用「另存为」保存副本。",
    );
    return false;
  }
  try {
    // 先标记自身写入：OS 可能在 invoke 返回前就投递监听事件
    markSelfWrite();
    if (doc.docType === "pdf" || isPdfPath(doc.filePath)) {
      if (!doc.pdfBase64) return false;
      // 1. 将高亮标注与便签元数据保存到持久化存储（不修改 PDF 二进制）
      await savePdfAnnotations(doc.filePath, {
        highlights: doc.pdfHighlights ?? [],
        notes: doc.pdfNotes ?? [],
        paperTheme: doc.pdfPaperTheme,
        currentPage: doc.pdfCurrentPage,
      });

      // 2. 仅在页面结构发生实质修改时才重新写磁盘文件（例如旋转/删除/重排等导致 pdfBase64 变化）
      const hasBinaryChange =
        Boolean(doc.savedPdfBase64) &&
        doc.pdfBase64 !== doc.savedPdfBase64;

      let modifiedAt = doc.modifiedAt;
      let newSize = doc.size;

      if (hasBinaryChange) {
        modifiedAt = await invoke<number>("write_binary_file", {
          path: doc.filePath,
          base64: doc.pdfBase64,
          encryptedHeader: doc.encryptedHeader,
        });
        newSize = (doc.encrypted ? 4096 : 0) + Math.floor((doc.pdfBase64.length * 3) / 4);
      }

      markSelfWrite();
      useAppStore.getState().patchDoc(id, {
        // 关键：绝对不改动 pdfBase64，保持当前内存中二进制引用不变，避免触发 PdfViewer 重载！
        savedPdfBase64: doc.pdfBase64,
        cleanPdfBase64: doc.pdfBase64,
        isDirty: false,
        modifiedAt,
        size: newSize,
      });
      void addRecentFile(doc.filePath);
      return true;
    }

    const modifiedAt = await invoke<number>("write_markdown_file", {
      path: doc.filePath,
      content: doc.content,
      encryptedHeader: doc.encryptedHeader,
      encoding: doc.encoding,
      eol: doc.eol,
    });
    markSelfWrite();
    // 写入期间用户可能继续输入：只有内容与写入时完全一致才能标记为已保存，
    // 否则会把「还没落盘的按键」当成已保存，自动保存与关闭确认都不再提示
    const latest = getDocById(id);
    const stillSame = (latest?.content ?? doc.content) === doc.content;
    useAppStore.getState().patchDoc(id, {
      savedContent: doc.content,
      isDirty: stillSame ? false : true,
      modifiedAt,
      // 加密文档落盘后会多出 4096 字节文件头
      size: (doc.encrypted ? 4096 : 0) + utf8Size(doc.content),
    });
    void addRecentFile(doc.filePath);
    return true;
  } catch (error) {
    await showMessage("保存失败", String(error));
    return false;
  }
}

export async function saveDocAs(id: string): Promise<boolean> {
  const doc = getDocById(id);
  if (!doc) return false;
  try {
    const isBlank = doc.docType === "blank" && !doc.filePath;
    const isPdf = doc.docType === "pdf" || isPdfPath(doc.filePath);
    const defaultPath = doc.filePath ?? (isPdf ? "未命名.pdf" : isBlank ? "未命名" : "未命名.md");
    const target = await invoke<string | null>("save_file_dialog", {
      defaultPath,
      filterAll: isBlank,
      filterPdf: isPdf,
    });
    if (!target) return false;
    markSelfWrite();

    if (isPdf) {
      if (!doc.pdfBase64) return false;
      const modifiedAt = await invoke<number>("write_binary_file", {
        path: target,
        base64: doc.pdfBase64,
        encryptedHeader: doc.encryptedHeader,
      });
      await savePdfAnnotations(target, {
        highlights: doc.pdfHighlights ?? [],
        notes: doc.pdfNotes ?? [],
        paperTheme: doc.pdfPaperTheme,
        currentPage: doc.pdfCurrentPage,
      });
      markSelfWrite();
      const oldPath = doc.filePath;
      useAppStore.getState().patchDoc(id, {
        filePath: target,
        docType: "pdf",
        savedPdfBase64: doc.pdfBase64,
        cleanPdfBase64: doc.pdfBase64,
        isDirty: false,
        readOnly: false,
        size: (doc.encrypted ? 4096 : 0) + Math.floor((doc.pdfBase64.length * 3) / 4),
        modifiedAt,
      });
      if (oldPath && !samePath(oldPath, target)) void unwatchFile(oldPath);
      void addRecentFile(target);
      void watchFile(target);
      return true;
    }

    const modifiedAt = await invoke<number>("write_markdown_file", {
      path: target,
      content: doc.content,
      encryptedHeader: doc.encryptedHeader,
      encoding: doc.encoding,
      eol: doc.eol,
    });
    markSelfWrite();
    const oldPath = doc.filePath;
    useAppStore.getState().patchDoc(id, {
      filePath: target,
      savedContent: doc.content,
      isDirty: false,
      readOnly: false,
      // 加密文档另存为后仍是加密文档（沿用原文件头）
      size: (doc.encrypted ? 4096 : 0) + utf8Size(doc.content),
      modifiedAt,
    });
    if (oldPath && !samePath(oldPath, target)) void unwatchFile(oldPath);
    void addRecentFile(target);
    void watchFile(target);
    return true;
  } catch (error) {
    await showMessage("保存失败", String(error));
    return false;
  }
}

export async function saveActive(): Promise<boolean> {
  const doc = getActiveDoc();
  if (!doc) return false;
  return saveDoc(doc.id);
}

export async function saveActiveAs(): Promise<boolean> {
  const doc = getActiveDoc();
  if (!doc) return false;
  return saveDocAs(doc.id);
}

/* ------------------------------ 关闭 / 新建 --------------------------- */

export async function closeDocWithConfirm(id: string): Promise<boolean> {
  const doc = getDocById(id);
  if (!doc) return true;
  if (!(await ensureNoDirty(doc))) return false;
  if (doc.filePath) {
    // 同一路径可能还开在另一栏（分屏同文档对照）：只有无人引用时才解除监听，
    // 否则剩下的那个标签页再也收不到外部变更通知
    const stillReferenced = useAppStore
      .getState()
      .docs.some((other) => other.id !== id && samePath(other.filePath, doc.filePath));
    if (!stillReferenced) void unwatchFile(doc.filePath);
  }
  // 释放该文档的 PDF 撤销栈与自动清洗计数（每个快照都持有一份 base64 副本）。
  // 用动态 import 保持 pdf-lib 留在按需加载的分包里，不拖累首屏体积
  if (doc.docType === "pdf" || isPdfPath(doc.filePath)) {
    void import("../components/PDF/pdfService")
      .then((mod) => mod.disposePdfDocState(id))
      .catch(() => undefined);
  }
  useAppStore.getState().closeDoc(id);
  return true;
}

/** 关闭除 keepId 之外的全部标签页（逐个走未保存确认，任何一步取消即中止） */
export async function closeOtherDocsWithConfirm(keepId: string): Promise<boolean> {
  const others = useAppStore.getState().docs.filter((doc) => doc.id !== keepId);
  for (const doc of others) {
    if (!(await closeDocWithConfirm(doc.id))) return false;
  }
  return true;
}

/** 关闭 keepId 右侧的标签页（同栏内按显示顺序，逐个确认） */
export async function closeRightDocsWithConfirm(keepId: string): Promise<boolean> {
  const state = useAppStore.getState();
  const keep = state.docs.find((doc) => doc.id === keepId);
  if (!keep) return true;
  const pane = keep.pane ?? 0;
  const paneDocs = state.docs.filter((doc) => (doc.pane ?? 0) === pane);
  const keepIndex = paneDocs.findIndex((doc) => doc.id === keepId);
  if (keepIndex < 0) return true;
  for (const doc of paneDocs.slice(keepIndex + 1)) {
    if (!(await closeDocWithConfirm(doc.id))) return false;
  }
  return true;
}

export async function closeAllDocsWithConfirm(): Promise<boolean> {
  const docs = [...useAppStore.getState().docs];
  for (const doc of docs) {
    if (!(await closeDocWithConfirm(doc.id))) return false;
  }
  return true;
}

export async function newDocument(
  type: "markdown" | "blank" = "markdown",
  targetPane?: 0 | 1,
): Promise<void> {
  if (!(await ensureNoDirty(getActiveDoc()))) return;
  const isMd = type === "markdown";
  const activePane =
    targetPane !== undefined ? targetPane : (useAppStore.getState().layout.activePane ?? 0);
  const doc = createDoc({
    content: isMd ? EMPTY_DOC_PLACEHOLDER : "",
    savedContent: isMd ? EMPTY_DOC_PLACEHOLDER : "",
    docType: type,
    pane: activePane,
  });
  useAppStore.getState().addDoc(doc);
  useAppStore.getState().setViewMode("source");
}

/** 丢弃本地修改，重新从磁盘加载 */
export async function reloadDocFromDisk(id: string): Promise<boolean> {
  const doc = getDocById(id);
  if (!doc?.filePath) return false;
  try {
    if (doc.docType === "pdf" || isPdfPath(doc.filePath)) {
      const payload = await invoke<BinaryPayload>("read_binary_file", {
        path: doc.filePath,
      });
      const encrypted = payload.encrypted ?? false;
      useAppStore.getState().patchDoc(id, {
        pdfBase64: payload.dataBase64,
        savedPdfBase64: payload.dataBase64,
        isDirty: false,
        encrypted,
        encryptedHeader: payload.encryptedHeader ?? null,
        modifiedAt: payload.modifiedAt,
        size: payload.size,
      });
      return true;
    }

    const payload = await invoke<FilePayload>("read_markdown_file", {
      path: doc.filePath,
    });
    const encrypted = payload.encrypted ?? false;
    useAppStore.getState().patchDoc(id, {
      content: payload.content,
      savedContent: payload.content,
      isDirty: false,
      encrypted,
      encryptedHeader: payload.encryptedHeader ?? null,
      encoding: payload.encoding ?? doc.encoding,
      eol: payload.eol ?? doc.eol,
      readOnly: payload.size > LARGE_FILE_BYTES,
      modifiedAt: payload.modifiedAt,
      size: payload.size,
    });
    return true;
  } catch (error) {
    console.error("重新加载失败", error);
    return false;
  }
}

/* ------------------------------ 外链 ------------------------------ */

export async function openExternal(url: string): Promise<void> {
  if (!/^(https?:|mailto:|tel:)/i.test(url)) return;
  try {
    await openUrl(url);
  } catch (error) {
    await showMessage("无法打开链接", `${url}\n\n${String(error)}`);
  }
}
