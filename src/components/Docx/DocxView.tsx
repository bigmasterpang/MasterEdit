/**
 * Word 文档（.docx）**只读**视图 —— 排版渲染器。
 *
 * 定位（见 `docs/plan-docx.md`）：**结构忠实，不做像素级还原**。段落 / 表格 / 图片 /
 * 编号 / 字体字号颜色 / 缩进对齐都按文档来，但不强行分页（只在原来的分页位置画淡色提示线）。
 * 编辑交给 Word / WPS —— 保留「用其它应用编辑」「选择打开方式…」两个入口。
 *
 * 组成：
 *   · 顶部：文件名、已解密标记、规模统计、复制 / 查找 / 大纲 / 缩放 / 打开方式；
 *   · `DocxBlocks`：虚拟滚动块列表（高度估算 + 测量修正）；
 *   · `DocxOutline`：大纲侧栏（按 `outlineLevel` 收集）；
 *   · `DocxFind`：`Ctrl+F` 查找面板；
 *   · 表格右键：复制为制表符文本（可直接粘进 Excel）。
 */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { useAppStore } from "../../stores/appStore";
import type { DocTable, DocumentInfo } from "../../types";
import { Icon } from "../common/Icon";
import { openActiveDocWithExternalApp } from "../../utils/fileActions";
import { formatBytes } from "../../utils/timing";
import { readMigratedItem, writeItem } from "../../utils/storage";
import { DocxBlocks, type DocxViewMode } from "./DocxBlocks";
import { DocxFindBar, useDocxFind } from "./DocxFind";
import { DocxOutline } from "./DocxOutline";
import { useDocxBlocks } from "./useDocxBlocks";
import {
  DEFAULT_PAGE_BG,
  DOCX_PAGE_BG_PRESETS,
  MEDIA_TIMEOUT_MS,
  normalizePageBg,
} from "./docxStyle";
import { copyTextToClipboard } from "./docxCopy";
import { blocksToPlainText, collectOutline, tableToTsv, type OutlineItem } from "./docxGrid";
import type { DocxRenderContext } from "./docxRender";

/** 缩放范围（Ctrl+滚轮 / 工具条按钮） */
const SCALE_MIN = 0.6;
const SCALE_MAX = 2.4;
const SCALE_STEP = 0.1;
/** 顶部提示条自动消失时间 */
const TOAST_MS = 2600;
/** Word 页面底色持久化键（默认白色 `#ffffff`） */
const PAGE_BG_STORAGE_KEY = "masteredit.docx.pageBg";
const PAGE_BG_LEGACY_KEY = "mastermd.docx.pageBg";

interface Toast {
  kind: "ok" | "error";
  text: string;
}

interface MenuState {
  x: number;
  y: number;
  table: DocTable;
}

export function DocxView({ docId }: { docId: string }) {
  const filePath = useAppStore((s) => s.docs.find((d) => d.id === docId)?.filePath ?? null);
  const modifiedAt = useAppStore((s) => s.docs.find((d) => d.id === docId)?.modifiedAt ?? 0);
  const active = useAppStore((s) => s.activeId === docId);
  const fileName = filePath ? (filePath.split(/[\\/]/).pop() ?? filePath) : "（未保存的文档）";

  const [info, setInfo] = useState<DocumentInfo | null>(null);
  const [infoError, setInfoError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [showXml, setShowXml] = useState(false);
  const [xml, setXml] = useState("");
  const [scale, setScale] = useState(1);
  /** 视图模式：**默认分页**（用户明确要"一页一页"），可切回连续流 */
  const [viewMode, setViewMode] = useState<DocxViewMode>("paged");
  /** 页面底色：**默认白色** `#ffffff`，可切换预设或自定义颜色并持久化 */
  const [pageBg, setPageBg] = useState<string>(() =>
    normalizePageBg(readMigratedItem(PAGE_BG_STORAGE_KEY, PAGE_BG_LEGACY_KEY) ?? DEFAULT_PAGE_BG),
  );
  const updatePageBg = useCallback((value: string) => {
    const next = normalizePageBg(value);
    setPageBg(next);
    writeItem(PAGE_BG_STORAGE_KEY, next);
  }, []);
  const [copyBusy, setCopyBusy] = useState(false);
  const [toast, setToast] = useState<Toast | null>(null);
  const [menu, setMenu] = useState<MenuState | null>(null);
  const [outlineOpen, setOutlineOpen] = useState(false);
  const [outline, setOutline] = useState<{ items: OutlineItem[]; partial: boolean } | null>(null);
  const [highlightBlock, setHighlightBlock] = useState<number | null>(null);
  const [scrollRequest, setScrollRequest] = useState<{ index: number; token: number } | null>(null);
  const scrollTokenRef = useRef(0);

  const api = useDocxBlocks(filePath, modifiedAt);

  /* ------------------ 文档信息（规模统计 / 解密标记） ------------------ */
  useEffect(() => {
    if (!filePath) return;
    let cancelled = false;
    setLoading(true);
    setInfoError(null);
    void (async () => {
      try {
        const result = await invoke<DocumentInfo>("document_info", { path: filePath });
        if (!cancelled) setInfo(result);
      } catch (reason) {
        if (!cancelled) setInfoError(String(reason));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [filePath, modifiedAt]);

  /* ------------------ 图片：按 (路径|media) 缓存 + 请求去重 ------------------ */
  const mediaCacheRef = useRef(new Map<string, string>());
  const mediaPendingRef = useRef(new Map<string, Promise<string>>());
  const lastMediaPathRef = useRef(filePath);
  if (lastMediaPathRef.current !== filePath) {
    // 换文档就清空上一个文档的图片缓存；原地 clear 保持 Map 引用不变，避免 renderContext 拿到旧 Map
    lastMediaPathRef.current = filePath;
    mediaCacheRef.current.clear();
    mediaPendingRef.current.clear();
  }
  const mediaCacheKey = useCallback(
    (media: string) => `${filePath ?? ""}|${media}`,
    [filePath],
  );
  const loadMedia = useCallback(
    (media: string): Promise<string> => {
      const key = `${filePath ?? ""}|${media}`;
      const cached = mediaCacheRef.current.get(key);
      if (cached) return Promise.resolve(cached);
      const pending = mediaPendingRef.current.get(key);
      if (pending) return pending; // 同一张图并发请求只发一次
      /**
       * **带超时**：后端 `document_media` 正常只要几毫秒（实测 2.4ms / 23KB）
       * 但异常文档上可能长时间不返回。若让它一直挂着，既会卡住这张图，
       * 也会让 `mediaPendingRef` 里那条"永远 pending"的记录挡住后续所有重试。
       * 超时即 reject（并清掉 pending 记录）→ 组件走失败态、下次进入视口重新发请求。
       */
      const task = new Promise<string>((resolve, reject) => {
        const timer = window.setTimeout(() => {
          mediaPendingRef.current.delete(key);
          reject(`图片读取超时（超过 ${MEDIA_TIMEOUT_MS / 1000} 秒）`);
        }, MEDIA_TIMEOUT_MS);
        invoke<string>("document_media", { path: filePath, media })
          .then((src) => {
            window.clearTimeout(timer);
            mediaCacheRef.current.set(key, src);
            mediaPendingRef.current.delete(key);
            resolve(src);
          })
          .catch((reason: unknown) => {
            window.clearTimeout(timer);
            mediaPendingRef.current.delete(key);
            reject(reason);
          });
      });
      mediaPendingRef.current.set(key, task);
      return task;
    },
    [filePath],
  );

  /* ------------------ 顶部提示条 ------------------ */
  const showToast = useCallback((kind: Toast["kind"], text: string) => {
    setToast({ kind, text });
  }, []);
  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(null), TOAST_MS);
    return () => window.clearTimeout(timer);
  }, [toast]);

  /* ------------------ 复制 ------------------ */
  const copyAllText = useCallback(async () => {
    setCopyBusy(true);
    try {
      const { blocks, complete } = await api.loadAll();
      const text = blocksToPlainText(blocks);
      const outcome = await copyTextToClipboard(text);
      showToast(
        outcome.ok ? "ok" : "error",
        complete ? outcome.message : `${outcome.message}（文档较大，只复制了前 ${blocks.length} 块）`,
      );
    } catch (reason) {
      showToast("error", `复制失败：${String(reason)}`);
    } finally {
      setCopyBusy(false);
    }
  }, [api, showToast]);

  const copyTable = useCallback(
    async (table: DocTable) => {
      const outcome = await copyTextToClipboard(tableToTsv(table));
      showToast(outcome.ok ? "ok" : "error", outcome.message);
    },
    [showToast],
  );

  /* ------------------ 查找 ------------------ */
  const jumpToBlock = useCallback((block: number) => {
    scrollTokenRef.current += 1;
    setHighlightBlock(block);
    setScrollRequest({ index: block, token: scrollTokenRef.current });
  }, []);
  const find = useDocxFind({ path: filePath, active, jump: (hit) => jumpToBlock(hit.block) });
  const findOpen = find.open;
  useEffect(() => {
    if (!findOpen) setHighlightBlock(null);
  }, [findOpen]);

  /* ------------------ 大纲：打开时把整篇块补齐再收集 ------------------ */
  const outlineLoading = api.loading || (outlineOpen && outline === null);
  useEffect(() => {
    if (!outlineOpen) return;
    let cancelled = false;
    void (async () => {
      const { blocks, complete } = await api.loadAll();
      if (cancelled) return;
      setOutline({ items: collectOutline(blocks), partial: !complete });
    })();
    return () => {
      cancelled = true;
    };
    // api.version：后台预取到新块后大纲要跟着长出来
  }, [outlineOpen, api, api.version]);

  /* ------------------ Ctrl+滚轮缩放 ------------------ */
  /**
   * 用**回调 ref**而不是 useEffect 挂原生监听：首屏是「正在解析…」分支，
   * 真正的根节点要等加载完才挂上；用 `useEffect(..., [])` 会在 ref 还是 null 时
   * 就跑完，监听器永远挂不上（Ctrl+滚轮缩放会完全失效）。
   * 必须是原生监听（React 的 onWheel 是被动监听，拿不到 preventDefault）。
   */
  const wheelDetachRef = useRef<(() => void) | null>(null);
  const attachRoot = useCallback((el: HTMLDivElement | null) => {
    wheelDetachRef.current?.();
    wheelDetachRef.current = null;
    if (!el) return;
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) return;
      // 拦住上层（DocView 的 Ctrl+滚轮是改 markdown 字号的，对 docx 没意义）
      event.preventDefault();
      event.stopPropagation();
      const delta = event.deltaY < 0 ? SCALE_STEP : -SCALE_STEP;
      setScale((value) => Math.min(SCALE_MAX, Math.max(SCALE_MIN, Number((value + delta).toFixed(2)))));
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    wheelDetachRef.current = () => el.removeEventListener("wheel", onWheel);
  }, []);

  /* ------------------ 右键菜单：点别处 / Esc / 滚动都关掉 ------------------ */
  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(null);
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") close();
    };
    window.addEventListener("mousedown", close);
    window.addEventListener("keydown", onKeyDown);
    return () => {
      window.removeEventListener("mousedown", close);
      window.removeEventListener("keydown", onKeyDown);
    };
  }, [menu]);

  /**
   * 文档滚动容器的 ref：图片懒加载的 IntersectionObserver 要拿它当 root
   * （`DocxBlocks` 在 layout 阶段把容器元素写进来）。
   */
  const scrollRootRef = useRef<HTMLDivElement | null>(null);

  const renderContext: DocxRenderContext = useMemo(
    () => ({
      scale,
      pageBg,
      mediaCache: mediaCacheRef.current,
      mediaCacheKey,
      loadMedia,
      highlightBlock,
      scrollRootRef,
      // 页码/总页数由 `DocxBlocks` 逐页覆盖（`PAGE`/`NUMPAGES` 域要用）；这里是缺省值
      pageNumber: 1,
      totalPages: 1,
      // 纸张可容纳块宽同样由 `DocxBlocks` 按纸张几何覆盖（0 = 退回版心宽）
      maxBlockWidth: 0,
      onTableContextMenu: (event, table) => setMenu({ x: event.clientX, y: event.clientY, table }),
    }),
    [scale, pageBg, mediaCacheKey, loadMedia, highlightBlock],
  );

  /* ------------------ 原始 document.xml（排查用，保留原能力） ------------------ */
  const toggleXml = useCallback(async () => {
    if (showXml) {
      setShowXml(false);
      return;
    }
    if (!filePath) return;
    if (!xml) {
      try {
        const text = await invoke<string>("document_xml", { path: filePath });
        setXml(text);
      } catch (reason) {
        showToast("error", String(reason));
        return;
      }
    }
    setShowXml(true);
  }, [filePath, showXml, xml, showToast]);

  if (loading) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-2 text-muted">
        <Icon name="loader" size={24} className="animate-spin text-accent" />
        <div className="text-[12px]">正在解析 Word 文档…</div>
      </div>
    );
  }

  if (infoError) {
    return (
      <div className="flex h-full items-center justify-center p-6">
        <div className="max-w-[520px] rounded-lg border border-line bg-panel p-4">
          <div className="mb-2 flex items-center gap-2 text-[13px] font-medium text-danger">
            <Icon name="alert-triangle" size={15} />
            打不开这个文档
          </div>
          <div className="whitespace-pre-wrap text-[12px] leading-relaxed text-muted">{infoError}</div>
          <button
            type="button"
            onClick={() => void openActiveDocWithExternalApp("default")}
            className="mt-3 rounded border border-line px-2 py-1 text-[12px] text-fg hover:bg-hover"
          >
            用系统默认程序打开
          </button>
        </div>
      </div>
    );
  }

  const encrypted = info?.encrypted ?? api.encrypted;
  const paragraphs = info?.paragraphs ?? 0;
  const tables = info?.tables ?? 0;
  const images = info?.images ?? 0;

  return (
    <div ref={attachRoot} className="flex h-full flex-col overflow-hidden bg-app">
      {/* ------------------------------ 顶部工具条 ------------------------------ */}
      <div className="border-b border-line bg-panel px-4 py-2">
        <div className="flex flex-wrap items-center gap-2 text-[12px]">
          <span
            data-docx-filename="true"
            className="max-w-[280px] truncate font-medium text-fg"
            title={filePath ?? undefined}
          >
            {fileName}
          </span>
          <span className="rounded bg-accent/10 px-1.5 py-0.5 font-medium text-accent">
            Word 文档（只读）
          </span>
          {encrypted ? (
            <span
              data-docx-encrypted="true"
              className="rounded bg-accent/10 px-1.5 py-0.5 text-accent"
              title="企业加密文档：已在内存中解密查看，不会写出明文副本"
            >
              已解密
            </span>
          ) : null}
          <span
            data-docx-stats="true"
            className="text-muted"
            title="「段落」是后端统计的全部段落（含表格单元格内的段落），「顶层块」是文档正文里的块数 —— 两者不是一个口径"
          >
            段落 {paragraphs} · 表格 {tables} · 图片 {images} · 顶层块 {api.total}
          </span>
          <span className="text-faint">
            包内 {info?.parts.length ?? 0} 个部件 ·{" "}
            {formatBytes((info?.parts ?? []).reduce((total, part) => total + part.size, 0))}
          </span>
        </div>

        <div className="mt-2 flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => void openActiveDocWithExternalApp("default")}
            className="flex items-center gap-1 rounded border border-line px-2 py-1 text-[12px] text-fg hover:bg-hover"
            title="交给系统里的 Word / WPS 编辑（加密文档由系统驱动透明解密）"
          >
            <Icon name="external-link" size={13} />
            用其它应用编辑
          </button>
          <button
            type="button"
            onClick={() => void openActiveDocWithExternalApp("choose")}
            className="rounded border border-line px-2 py-1 text-[12px] text-fg hover:bg-hover"
          >
            选择打开方式…
          </button>

          <span className="mx-1 h-4 w-px bg-line" />

          <button
            type="button"
            data-docx-copy-all="true"
            disabled={copyBusy}
            onClick={() => void copyAllText()}
            className="flex items-center gap-1 rounded border border-line px-2 py-1 text-[12px] text-fg hover:bg-hover disabled:opacity-50"
            title="把整篇文档的文字复制到剪贴板（表格按制表符展开）"
          >
            <Icon name="copy" size={13} />
            {copyBusy ? "复制中…" : "复制全部文本"}
          </button>
          <button
            type="button"
            data-docx-find-toggle="true"
            onClick={() => (find.open ? find.close() : find.openPanel())}
            className={`flex items-center gap-1 rounded border px-2 py-1 text-[12px] hover:bg-hover ${
              find.open ? "border-accent text-accent" : "border-line text-fg"
            }`}
            title="查找（Ctrl+F）"
          >
            <Icon name="search" size={13} />
            查找
          </button>
          <button
            type="button"
            data-docx-outline-toggle="true"
            onClick={() => setOutlineOpen((value) => !value)}
            className={`flex items-center gap-1 rounded border px-2 py-1 text-[12px] hover:bg-hover ${
              outlineOpen ? "border-accent text-accent" : "border-line text-fg"
            }`}
            title="按大纲级别（标题）生成目录"
          >
            <Icon name="list" size={13} />
            大纲
          </button>

          <span className="mx-1 h-4 w-px bg-line" />

          {/* 分页 / 连续：默认分页（用户要的是"一页一页"），连续模式保留原来的一条流 */}
          <div
            data-docx-mode-switch="true"
            className="flex items-center overflow-hidden rounded border border-line"
            title="分页：按纸张一页一页显示（默认）；连续：一条流滚到底"
          >
            <button
              type="button"
              data-docx-mode-paged="true"
              aria-pressed={viewMode === "paged"}
              onClick={() => setViewMode("paged")}
              className={`px-2 py-1 text-[12px] ${
                viewMode === "paged" ? "bg-accent/15 text-accent" : "text-muted hover:bg-hover"
              }`}
            >
              分页
            </button>
            <span className="h-4 w-px bg-line" />
            <button
              type="button"
              data-docx-mode-continuous="true"
              aria-pressed={viewMode === "continuous"}
              onClick={() => setViewMode("continuous")}
              className={`px-2 py-1 text-[12px] ${
                viewMode === "continuous" ? "bg-accent/15 text-accent" : "text-muted hover:bg-hover"
              }`}
            >
              连续
            </button>
          </div>

          <span className="mx-1 h-4 w-px bg-line" />

          <div className="flex items-center gap-1">
            <button
              type="button"
              data-docx-zoom-out="true"
              onClick={() => setScale((value) => Math.max(SCALE_MIN, Number((value - SCALE_STEP).toFixed(2))))}
              className="flex h-6 w-6 items-center justify-center rounded border border-line text-muted hover:bg-hover"
              title="缩小（Ctrl+滚轮）"
            >
              <Icon name="zoom-out" size={13} />
            </button>
            <span data-docx-zoom-value="true" className="w-[42px] text-center text-[11px] text-muted">
              {Math.round(scale * 100)}%
            </span>
            <button
              type="button"
              data-docx-zoom-in="true"
              onClick={() => setScale((value) => Math.min(SCALE_MAX, Number((value + SCALE_STEP).toFixed(2))))}
              className="flex h-6 w-6 items-center justify-center rounded border border-line text-muted hover:bg-hover"
              title="放大（Ctrl+滚轮）"
            >
              <Icon name="zoom-in" size={13} />
            </button>
          </div>

          <span className="mx-1 h-4 w-px bg-line" />

          {/* 页面底色：默认白色（#ffffff），支持一键切换预设或自定义拾色 */}
          <div
            data-docx-bg-picker="true"
            className="flex items-center gap-1.5"
            title="调整 Word 页面底色（默认白色）"
          >
            <span className="text-[11px] text-muted">底色</span>
            {DOCX_PAGE_BG_PRESETS.map((preset) => {
              const activePreset = pageBg === preset.color;
              return (
                <button
                  key={preset.id}
                  type="button"
                  data-docx-bg-preset={preset.id}
                  data-docx-bg-color={preset.color}
                  aria-pressed={activePreset}
                  onClick={() => updatePageBg(preset.color)}
                  title={`底色：${preset.label}（${preset.color}）`}
                  className={`h-5 w-5 rounded border transition-transform ${
                    activePreset
                      ? "scale-105 border-accent ring-1 ring-accent"
                      : "border-line hover:scale-105"
                  }`}
                  style={{ backgroundColor: preset.color }}
                />
              );
            })}
            <label
              className="relative flex h-5 w-5 cursor-pointer items-center justify-center rounded border border-line text-[10px] text-muted hover:bg-hover"
              title={`自定义底色（当前 ${pageBg}）`}
            >
              🎨
              <input
                type="color"
                data-docx-bg-custom="true"
                value={pageBg}
                onChange={(event) => updatePageBg(event.target.value)}
                className="sr-only"
              />
            </label>
          </div>

          <button
            type="button"
            onClick={() => void toggleXml()}
            className="rounded border border-line px-2 py-1 text-[12px] text-muted hover:bg-hover hover:text-fg"
          >
            {showXml ? "收起原始 XML" : "查看原始 XML"}
          </button>
        </div>
      </div>

      {/* ------------------------------ 正文 ------------------------------ */}
      {showXml ? (
        <div className="min-h-0 flex-1 overflow-auto p-4">
          <pre className="whitespace-pre-wrap break-all rounded border border-line bg-panel p-3 font-mono text-[11px] leading-relaxed text-muted">
            {xml.slice(0, 400000)}
            {xml.length > 400000 ? "\n\n…（内容过长，仅显示前 400KB）" : ""}
          </pre>
        </div>
      ) : (
        <div className="flex min-h-0 flex-1">
          {outlineOpen ? (
            <DocxOutline
              items={outline?.items ?? []}
              loading={outlineLoading}
              partial={outline?.partial ?? false}
              onJump={jumpToBlock}
            />
          ) : null}
          <div className="relative min-w-0 flex-1">
            <DocxBlocks api={api} ctx={renderContext} mode={viewMode} scrollRequest={scrollRequest} />
            <DocxFindBar api={find} />

            {api.error ? (
              <div className="absolute bottom-3 left-1/2 z-30 -translate-x-1/2 rounded-lg border border-danger/40 bg-elevated/95 px-3 py-2 text-[12px] text-fg shadow-lg">
                <span className="text-danger">取块失败：</span>
                <span className="text-muted">{api.error}</span>
                <button
                  type="button"
                  onClick={() => api.retry()}
                  className="ml-2 rounded border border-line px-1.5 py-0.5 text-[11px] hover:bg-hover"
                >
                  重试
                </button>
              </div>
            ) : null}

            {toast ? (
              <div
                data-docx-toast="true"
                data-docx-toast-kind={toast.kind}
                className={`absolute bottom-3 left-1/2 z-40 -translate-x-1/2 rounded-lg border px-3 py-1.5 text-[12px] shadow-lg ${
                  toast.kind === "ok"
                    ? "border-line bg-elevated/95 text-fg"
                    : "border-danger/50 bg-elevated/95 text-danger"
                }`}
              >
                {toast.text}
              </div>
            ) : null}
          </div>
        </div>
      )}

      {/* ------------------------------ 表格右键菜单 ------------------------------ */}
      {menu ? (
        <div
          data-docx-table-menu="true"
          className="fixed z-50 min-w-[188px] rounded-lg border border-line bg-elevated py-1 shadow-xl"
          style={{ left: `${menu.x}px`, top: `${menu.y}px` }}
          onMouseDown={(event) => event.stopPropagation()}
        >
          <button
            type="button"
            data-docx-table-copy-tsv="true"
            onClick={() => {
              setMenu(null);
              void copyTable(menu.table);
            }}
            className="flex w-full items-center gap-2 px-3 py-1.5 text-left text-[12px] text-fg hover:bg-hover"
          >
            <Icon name="copy" size={13} className="text-muted" />
            复制为制表符文本（可粘贴到 Excel）
          </button>
          <button
            type="button"
            onClick={() => {
              setMenu(null);
              void copyAllText();
            }}
            className="flex w-full items-center gap-2 px-3 py-1.5 text-left text-[12px] text-fg hover:bg-hover"
          >
            <Icon name="file-text" size={13} className="text-muted" />
            复制整篇文档文本
          </button>
        </div>
      ) : null}
    </div>
  );
}
