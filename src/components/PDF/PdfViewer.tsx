import { useEffect, useRef, useState, useCallback, useMemo } from "react";
import * as pdfjsLib from "pdfjs-dist";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.js?url";
import "pdfjs-dist/web/pdf_viewer.css";
import { Icon } from "../common/Icon";
import { ContextMenu, type ContextMenuItem } from "../common/ContextMenu";
import { useAppStore } from "../../stores/appStore";
import { useSettingsStore } from "../../stores/settingsStore";
import { askPdfPassword, showMessage } from "../../stores/dialogStore";
import { fileName } from "../../utils/filePath";
import type { PdfHighlight, PdfNote } from "../../types";
import {
  addPdfHighlight,
  addPdfNote,
  base64ToBytes,
  cleanBurnedHighlightsFromPdf,
  clearAllHighlights,
  clearPageHighlights,
  copyCanvasToClipboard,
  deletePdfNote,
  deletePdfPage,
  extractPdfPage,
  PDF_PAPER_THEMES,
  recordAutoCleanedCount,
  redoPdf,
  registerPdfDocument,
  removePdfHighlight,
  repairBurnedPdfDocument,
  rotateAllPdfPages,
  rotatePdfPage,
  undoPdf,
  unregisterPdfDocument,
  updatePdfHighlight,
  updatePdfNote,
  type OutlineItem,
} from "./pdfService";

// 设置 PDF.js Worker 路径
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

interface PdfViewerProps {
  docId: string;
  pane?: 0 | 1;
  isDark: boolean;
}

interface PdfSelectionMenuState {
  x: number;
  y: number;
  selectedText: string;
  pageNum: number;
  clientRects: DOMRect[];
  pageRect: DOMRect | null;
}

interface HighlightMenuState {
  x: number;
  y: number;
  highlightId: string;
  pageNum: number;
}

interface PageContextMenuState {
  x: number;
  y: number;
  pageNum: number;
  xPercent: number;
  yPercent: number;
}

const SCALE_PRESETS = [
  { label: "50%", value: 0.5 },
  { label: "75%", value: 0.75 },
  { label: "100%", value: 1.0 },
  { label: "125%", value: 1.25 },
  { label: "150%", value: 1.5 },
  { label: "200%", value: 2.0 },
  { label: "300%", value: 3.0 },
];

export function PdfViewer({ docId, pane, isDark }: PdfViewerProps) {
  const doc = useAppStore((s) => s.docs.find((d) => d.id === docId) ?? null);
  const activePane = useAppStore((s) => s.layout.activePane);
  const isPaneActive = pane === undefined || pane === activePane;
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [pdfProxy, setPdfProxy] = useState<pdfjsLib.PDFDocumentProxy | null>(null);
  const [numPages, setNumPages] = useState(0);
  const [currentPage, setCurrentPage] = useState(1);
  const [scale, setScale] = useState<number>(1.2);
  const [fitMode, setFitMode] = useState<"custom" | "width" | "page">("width");

  // pdfjs 文档与加载任务必须持有引用：不释放的话每次重载都会残留一个 Web Worker
  // 与整份已解析文档（旋转 / 删除 / 撤销 / 重排都会触发重载）
  const loadingTaskRef = useRef<pdfjsLib.PDFDocumentLoadingTask | null>(null);
  const pdfProxyRef = useRef<pdfjsLib.PDFDocumentProxy | null>(null);
  /** 加载代次：切换文档或卸载后，被取代的那次加载不得再注册文档 */
  const loadRunRef = useRef(0);
  /** 密码弹窗里输入的密码：等加载成功后再落库 */
  const pendingPasswordRef = useRef<string | null>(null);

  // 展开的批注/便签状态
  const [activeNoteId, setActiveNoteId] = useState<string | null>(null);
  const [activeCommentHlId, setActiveCommentHlId] = useState<string | null>(null);

  // 浮层菜单位置状态（采用 fixed 坐标，避免被工具栏截断）
  const [themeMenuPos, setThemeMenuPos] = useState<{ left: number; top: number } | null>(null);
  const [scaleMenuPos, setScaleMenuPos] = useState<{ left: number; top: number } | null>(null);

  // 右键菜单状态
  const [selectionMenu, setSelectionMenu] = useState<PdfSelectionMenuState | null>(null);
  const [highlightMenu, setHighlightMenu] = useState<HighlightMenuState | null>(null);
  const [pageContextMenu, setPageContextMenu] = useState<PageContextMenuState | null>(null);

  const containerRef = useRef<HTMLDivElement>(null);
  const pageRefs = useRef<Map<number, HTMLDivElement>>(new Map());
  const themeBtnRef = useRef<HTMLButtonElement>(null);
  const themeDropdownRef = useRef<HTMLDivElement>(null);
  const scaleBtnRef = useRef<HTMLButtonElement>(null);
  const scaleDropdownRef = useRef<HTMLDivElement>(null);

  // 重载后保持当前页锚定
  const targetPageAfterReload = useRef<number | null>(null);
  // 记录已成功加载的 Base64 内容，避免相同内容重复重载导致滚动位置重置
  const loadedBase64Ref = useRef<string>("");

  const pdfBase64 = doc?.pdfBase64 ?? "";
  const docName = doc?.filePath ? fileName(doc.filePath) : "PDF 文档";

  // 当前底色主题
  const paperTheme = doc?.pdfPaperTheme || "white";
  const activePaperTheme = useMemo(
    () => PDF_PAPER_THEMES.find((t) => t.id === paperTheme) || PDF_PAPER_THEMES[0],
    [paperTheme],
  );

  // 监听来自侧边栏的聚焦标注事件
  useEffect(() => {
    const handler = (e: any) => {
      if (e.detail?.docId === docId) {
        if (e.detail.type === "note") {
          setActiveNoteId(e.detail.annotationId);
          setActiveCommentHlId(null);
        } else if (e.detail.type === "highlight") {
          setActiveCommentHlId(e.detail.annotationId);
          setActiveNoteId(null);
        }
      }
    };
    window.addEventListener("pdf-focus-annotation" as any, handler);
    return () => window.removeEventListener("pdf-focus-annotation" as any, handler);
  }, [docId]);

  // 点击外部关闭弹出层菜单
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      const target = e.target as Node;
      if (
        themeDropdownRef.current &&
        !themeDropdownRef.current.contains(target) &&
        !themeBtnRef.current?.contains(target)
      ) {
        setThemeMenuPos(null);
      }
      if (
        scaleDropdownRef.current &&
        !scaleDropdownRef.current.contains(target) &&
        !scaleBtnRef.current?.contains(target)
      ) {
        setScaleMenuPos(null);
      }
    };
    window.addEventListener("mousedown", handleClickOutside);
    return () => window.removeEventListener("mousedown", handleClickOutside);
  }, []);

  // 加载 PDF 文档
  const loadPdf = useCallback(
    async (password?: string) => {
      if (!pdfBase64) return;
      // 关键守卫：如果相同的二进制内容已经加载过，绝对不重新 loadPdf，防止页面重新卸载导致滚动跳顶
      if (loadedBase64Ref.current === pdfBase64 && pdfProxy) {
        return;
      }
      // 本次加载的代次：被新的加载取代或组件卸载后用它判定是否还该继续
      const runId = ++loadRunRef.current;
      setLoading(true);
      setError(null);

      try {
        // 自动检测并剥离旧版残留的物理高亮涂层，打开即恢复纯净
        let effectiveBase64 = pdfBase64;
        const { cleanedBase64, removedCount } = await cleanBurnedHighlightsFromPdf(pdfBase64);
        if (removedCount > 0 && cleanedBase64 !== pdfBase64) {
          effectiveBase64 = cleanedBase64;
          recordAutoCleanedCount(docId, removedCount);
          loadedBase64Ref.current = cleanedBase64;
          useAppStore.getState().patchDoc(docId, {
            pdfBase64: cleanedBase64,
            cleanPdfBase64: cleanedBase64,
            isDirty: true,
          });
        }

        const rawBytes = base64ToBytes(effectiveBase64);
        const loadingTask = pdfjsLib.getDocument({
          data: rawBytes,
          password: password ?? doc?.pdfPassword,
          cMapUrl: "https://unpkg.com/pdfjs-dist@3.11.174/cmaps/",
          cMapPacked: true,
        });

        // 密码回调处理
        loadingTask.onPassword = async (updatePassword: (pw: string) => void, reason: number) => {
          const errMsg = reason === 2 ? "密码不正确，请重新输入" : undefined;
          const pw = await askPdfPassword(docName, errMsg);
          if (pw !== null) {
            updatePassword(pw);
            // 只暂存，不 patchDoc：pdfPassword 是 loadPdf 的依赖，
            // 在这里写库会立刻重启一次加载，并让上一次的密码弹窗等待一个永不结算的 Promise
            pendingPasswordRef.current = pw;
          } else {
            setError("需要密码才能查看此 PDF 文档。");
            setLoading(false);
          }
        };

        const proxy = await loadingTask.promise;

        // 本次加载是否已被更新的加载/卸载取代：是则释放资源后直接退出，
        // 否则会注册一个没有人负责回收的文档代理
        if (loadRunRef.current !== runId) {
          void proxy.destroy().catch(() => undefined);
          return;
        }

        // 释放上一份解析结果（留出渲染切换时间，避免正在绘制的旧页面引用失效）
        const previousProxy = pdfProxyRef.current;
        if (previousProxy && previousProxy !== proxy) {
          window.setTimeout(() => {
            void previousProxy.destroy().catch(() => undefined);
          }, 500);
        }
        pdfProxyRef.current = proxy;
        loadingTaskRef.current = loadingTask;
        setPdfProxy(proxy);
        setNumPages(proxy.numPages);
        loadedBase64Ref.current = effectiveBase64;

        // 关键修复：重载或旋转时保留用户当前的页码，而不是重置到第 1 页
        const existingDoc = useAppStore.getState().docs.find((d) => d.id === docId);
        const keepPage =
          targetPageAfterReload.current || existingDoc?.pdfCurrentPage || currentPage || 1;
        const safePage = Math.min(proxy.numPages, Math.max(1, keepPage));
        setCurrentPage(safePage);
        targetPageAfterReload.current = safePage;

        // 密码加载成功后再落库，供本会话内的后续重载直接使用
        if (pendingPasswordRef.current) {
          const pw = pendingPasswordRef.current;
          pendingPasswordRef.current = null;
          useAppStore.getState().patchDoc(docId, { pdfPassword: pw });
        }

        useAppStore.getState().patchDoc(docId, {
          pdfTotalPages: proxy.numPages,
          pdfCurrentPage: safePage,
        });

        // 递归提取完整多级大纲书签树（支持 Adobe Acrobat Pro DC 级别的多层子章节展示）
        const resolveOutlinePage = async (dest: any): Promise<number | undefined> => {
          if (!dest) return undefined;
          try {
            let destRef: any = dest;
            if (typeof destRef === "string") {
              destRef = await proxy.getDestination(destRef);
            }
            if (Array.isArray(destRef) && destRef.length > 0 && destRef[0] != null) {
              const target = destRef[0];
              if (typeof target === "number") {
                return target + 1;
              }
              if (typeof target === "object") {
                return (await proxy.getPageIndex(target)) + 1;
              }
            }
          } catch {
            /* ignore */
          }
          return undefined;
        };

        const parseOutlineNodes = async (nodes: any[]): Promise<OutlineItem[]> => {
          const result: OutlineItem[] = [];
          for (const item of nodes) {
            const pageIdx = await resolveOutlinePage(item.dest);
            const children =
              Array.isArray(item.items) && item.items.length > 0
                ? await parseOutlineNodes(item.items)
                : undefined;
            result.push({
              title: (item.title || "").replace(/\r?\n/g, " ").trim() || "未命名书签",
              dest: item.dest,
              pageIndex: pageIdx ?? children?.[0]?.pageIndex,
              items: children,
            });
          }
          return result;
        };

        let parsedOutline: OutlineItem[] = [];
        try {
          const rawOutline = await proxy.getOutline();
          if (rawOutline && rawOutline.length > 0) {
            parsedOutline = await parseOutlineNodes(rawOutline);
          }
        } catch {
          parsedOutline = [];
        }

        // 注册到全局共享服务
        registerPdfDocument(docId, proxy, parsedOutline);
        setLoading(false);
      } catch (err: any) {
        console.error("加载 PDF 失败", err);
        if (err?.name === "PasswordException") {
          setError("需要密码才能查看此 PDF 文档。");
        } else {
          setError(`加载 PDF 失败: ${err?.message || String(err)}`);
        }
        setLoading(false);
      }
    },
    [pdfBase64, doc?.pdfPassword, docName, docId],
  );

  useEffect(() => {
    void loadPdf();
    return () => {
      // 使正在进行的加载失效，并释放文档代理（含 Web Worker）
      loadRunRef.current += 1;
      unregisterPdfDocument(docId);
      loadedBase64Ref.current = "";
      const proxy = pdfProxyRef.current;
      pdfProxyRef.current = null;
      loadingTaskRef.current = null;
      if (proxy) {
        void proxy.destroy().catch(() => undefined);
      }
    };
  }, [loadPdf, docId]);

  // 跳转到指定页面（直接、瞬时直达，避免冗长缓动翻页）
  const scrollToPage = useCallback(
    (pageNum: number) => {
      const safePage = Math.max(1, Math.min(numPages, pageNum));
      const targetEl = pageRefs.current.get(safePage);
      const container = containerRef.current;
      if (targetEl && container) {
        const containerRect = container.getBoundingClientRect();
        const targetRect = targetEl.getBoundingClientRect();
        container.scrollTop += targetRect.top - containerRect.top - 16;
        setCurrentPage(safePage);
        useAppStore.getState().patchDoc(docId, { pdfCurrentPage: safePage });
      }
    },
    [numPages, docId],
  );

  // 旋转或重载后根据锚点重新定位回目标页
  useEffect(() => {
    if (!pdfProxy) return;
    const targetPage = targetPageAfterReload.current;
    if (targetPage && targetPage >= 1) {
      const timer1 = setTimeout(() => {
        scrollToPage(targetPage);
      }, 60);
      const timer2 = setTimeout(() => {
        scrollToPage(targetPage);
        targetPageAfterReload.current = null;
      }, 200);
      return () => {
        clearTimeout(timer1);
        clearTimeout(timer2);
      };
    }
  }, [pdfProxy, scrollToPage]);

  // 适应页面宽度计算
  const updateFitWidth = useCallback(async () => {
    if (!pdfProxy || !containerRef.current) return;
    try {
      const page = await pdfProxy.getPage(currentPage);
      const viewport = page.getViewport({ scale: 1.0 });
      const containerWidth = containerRef.current.clientWidth - 48;
      if (containerWidth > 0 && viewport.width > 0) {
        const newScale = Math.max(0.3, Math.min(3.5, containerWidth / viewport.width));
        setScale(newScale);
        setFitMode("width");
        useAppStore.getState().patchDoc(docId, { pdfScale: "width" });
      }
    } catch {
      /* ignore */
    }
  }, [pdfProxy, currentPage, docId]);

  // 适应整页高度计算
  const updateFitPage = useCallback(async () => {
    if (!pdfProxy || !containerRef.current) return;
    try {
      const page = await pdfProxy.getPage(currentPage);
      const viewport = page.getViewport({ scale: 1.0 });
      const containerHeight = containerRef.current.clientHeight - 48;
      if (containerHeight > 0 && viewport.height > 0) {
        const newScale = Math.max(0.3, Math.min(3.5, containerHeight / viewport.height));
        setScale(newScale);
        setFitMode("page");
        useAppStore.getState().patchDoc(docId, { pdfScale: "page" });
      }
    } catch {
      /* ignore */
    }
  }, [pdfProxy, currentPage, docId]);

  // 依赖里必须包含 pdfProxy：文档加载完成前 updateFitWidth 会直接 return，
  // 旧实现只在 fitMode 变化时触发，导致首屏永远停在默认 120%，
  // 明明工具栏高亮着「页宽」却要用户再点一次才生效
  useEffect(() => {
    if (!pdfProxy) return;
    if (fitMode === "width") {
      void updateFitWidth();
    } else if (fitMode === "page") {
      void updateFitPage();
    }
  }, [pdfProxy, fitMode]);

  // 视口滚动监听以追踪当前页码
  const handleScroll = () => {
    if (!containerRef.current || pageRefs.current.size === 0) return;
    const containerTop = containerRef.current.getBoundingClientRect().top;
    let closestPage = currentPage;
    let minDistance = Infinity;

    pageRefs.current.forEach((el, pNum) => {
      const rect = el.getBoundingClientRect();
      const distance = Math.abs(rect.top - containerTop - 40);
      if (distance < minDistance) {
        minDistance = distance;
        closestPage = pNum;
      }
    });

    if (closestPage !== currentPage) {
      setCurrentPage(closestPage);
      useAppStore.getState().patchDoc(docId, { pdfCurrentPage: closestPage });
    }
  };

  // 监听来自全局侧边栏的跳转事件
  useEffect(() => {
    const handler = (e: any) => {
      if (e.detail?.docId === docId && typeof e.detail?.pageNum === "number") {
        scrollToPage(e.detail.pageNum);
      }
    };
    window.addEventListener("pdf-jump-to-page" as any, handler);
    return () => window.removeEventListener("pdf-jump-to-page" as any, handler);
  }, [docId, scrollToPage]);

  // 支持 Ctrl + 滚轮一体化缩放（仅作用于当前 PDF 视口，不影响其它分栏）
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const handleWheel = (e: WheelEvent) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        e.stopPropagation();
        const delta = e.deltaY < 0 ? 0.15 : -0.15;
        setFitMode("custom");
        setScale((prev) => Math.max(0.3, Math.min(4.0, Number((prev + delta).toFixed(2)))));
      }
    };

    container.addEventListener("wheel", handleWheel, { passive: false });
    return () => {
      container.removeEventListener("wheel", handleWheel);
    };
  }, []);

  // 支持 PDF 操作撤销与重做快捷键 (Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z)
  useEffect(() => {
    if (!isPaneActive) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      // 避免干扰输入框或文本域中的正常撤回
      const target = e.target as HTMLElement | null;
      if (
        target &&
        (target.tagName === "INPUT" ||
          target.tagName === "TEXTAREA" ||
          target.isContentEditable)
      ) {
        return;
      }

      if (e.ctrlKey || e.metaKey) {
        if (e.key === "z" || e.key === "Z") {
          e.preventDefault();
          e.stopPropagation();
          targetPageAfterReload.current = currentPage;
          if (e.shiftKey) {
            void redoPdf(docId);
          } else {
            void undoPdf(docId);
          }
        } else if (e.key === "y" || e.key === "Y") {
          e.preventDefault();
          e.stopPropagation();
          targetPageAfterReload.current = currentPage;
          void redoPdf(docId);
        }
      }
    };

    window.addEventListener("keydown", handleKeyDown, true);
    return () => window.removeEventListener("keydown", handleKeyDown, true);
  }, [docId, currentPage, isPaneActive]);

  // 右键菜单智能分流：选中文本 vs 空白非内容区域
  const handleContextMenu = (e: React.MouseEvent) => {
    const selection = window.getSelection();
    const text = selection?.toString()?.trim() || "";
    const range = selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;

    let targetPage = currentPage;
    let pageEl: HTMLElement | null = null;
    let cur: HTMLElement | null = e.target as HTMLElement;

    while (cur && cur !== containerRef.current) {
      if (cur.dataset?.pageNumber) {
        targetPage = parseInt(cur.dataset.pageNumber, 10);
        pageEl = cur;
        break;
      }
      cur = cur.parentElement;
    }

    const clientRects = range ? Array.from(range.getClientRects()) : [];
    const pageRect = pageEl ? pageEl.getBoundingClientRect() : null;

    if (text.length > 0 && clientRects.length > 0) {
      // 1. 划词选区模式：弹出复制、高亮并添加注释、高亮菜单
      e.preventDefault();
      setPageContextMenu(null);
      setSelectionMenu({
        x: e.clientX,
        y: e.clientY,
        selectedText: text,
        pageNum: targetPage,
        clientRects,
        pageRect,
      });
    } else {
      // 2. 空白或非文本区域：弹出页面实用功能菜单（添加便签、复制页面为图片、另存、旋转等）
      e.preventDefault();
      window.getSelection()?.removeAllRanges();
      setSelectionMenu(null);

      let xPercent = 50;
      let yPercent = 50;
      if (pageRect && pageRect.width > 0 && pageRect.height > 0) {
        xPercent = Math.max(2, Math.min(95, ((e.clientX - pageRect.left) / pageRect.width) * 100));
        yPercent = Math.max(2, Math.min(95, ((e.clientY - pageRect.top) / pageRect.height) * 100));
      }

      setPageContextMenu({
        x: e.clientX,
        y: e.clientY,
        pageNum: targetPage,
        xPercent,
        yPercent,
      });
    }
  };

  // 划词选区右键菜单：支持先添加高亮，再针对高亮进行注释
  const selectionMenuGroups = useMemo<ContextMenuItem[][]>(() => {
    if (!selectionMenu) return [];
    const pageHls = (doc?.pdfHighlights ?? []).filter((h) => h.page === selectionMenu.pageNum);
    const hasAnyHls = (doc?.pdfHighlights ?? []).length > 0;

    return [
      [
        {
          label: "复制文本",
          hint: "Ctrl+C",
          icon: "copy",
          onClick: async () => {
            if (selectionMenu.selectedText) {
              await navigator.clipboard.writeText(selectionMenu.selectedText);
            }
            setSelectionMenu(null);
          },
        },
        {
          label: "高亮并添加注释",
          icon: "message-square" as any,
          onClick: () => {
            if (selectionMenu.pageRect && selectionMenu.clientRects.length > 0) {
              const newHl = addPdfHighlight(
                docId,
                selectionMenu.pageNum,
                selectionMenu.clientRects,
                selectionMenu.pageRect,
                "yellow",
                selectionMenu.selectedText,
                " ", // 初始化非空注释，触发注释卡片展现
              );
              if (newHl) {
                setActiveCommentHlId(newHl.id);
                setActiveNoteId(null);
              }
              window.getSelection()?.removeAllRanges();
            }
            setSelectionMenu(null);
          },
        },
      ],
      [
        {
          label: useSettingsStore.getState().colorblindMode
            ? "高亮标记 (明黄 · 实线)"
            : "高亮标记 (黄色)",
          icon: "bold",
          onClick: () => {
            if (selectionMenu.pageRect && selectionMenu.clientRects.length > 0) {
              addPdfHighlight(
                docId,
                selectionMenu.pageNum,
                selectionMenu.clientRects,
                selectionMenu.pageRect,
                "yellow",
                selectionMenu.selectedText,
              );
              window.getSelection()?.removeAllRanges();
            }
            setSelectionMenu(null);
          },
        },
        {
          label: useSettingsStore.getState().colorblindMode
            ? "高亮标记 (天青蓝 · 虚线)"
            : "高亮标记 (绿色)",
          onClick: () => {
            if (selectionMenu.pageRect && selectionMenu.clientRects.length > 0) {
              addPdfHighlight(
                docId,
                selectionMenu.pageNum,
                selectionMenu.clientRects,
                selectionMenu.pageRect,
                "green",
                selectionMenu.selectedText,
              );
              window.getSelection()?.removeAllRanges();
            }
            setSelectionMenu(null);
          },
        },
        {
          label: useSettingsStore.getState().colorblindMode
            ? "高亮标记 (洋红紫 · 双线)"
            : "高亮标记 (粉色)",
          onClick: () => {
            if (selectionMenu.pageRect && selectionMenu.clientRects.length > 0) {
              addPdfHighlight(
                docId,
                selectionMenu.pageNum,
                selectionMenu.clientRects,
                selectionMenu.pageRect,
                "pink",
                selectionMenu.selectedText,
              );
              window.getSelection()?.removeAllRanges();
            }
            setSelectionMenu(null);
          },
        },
      ],
      ...(pageHls.length > 0 || hasAnyHls
        ? [
            [
              ...(pageHls.length > 0
                ? [
                    {
                      label: "清除本页所有高亮",
                      icon: "trash" as any,
                      onClick: () => {
                        clearPageHighlights(docId, selectionMenu.pageNum);
                        setSelectionMenu(null);
                      },
                    },
                  ]
                : []),
              {
                label: "清除全部高亮",
                icon: "trash" as any,
                onClick: () => {
                  clearAllHighlights(docId);
                  setSelectionMenu(null);
                },
              },
            ],
          ]
        : []),
      [
        {
          label: "取消选区",
          icon: "x",
          onClick: () => {
            window.getSelection()?.removeAllRanges();
            setSelectionMenu(null);
          },
        },
      ],
    ];
  }, [selectionMenu, docId, doc?.pdfHighlights]);

  // 高亮项自身右键菜单（支持为高亮添加/编辑注释、切换颜色与删除）
  const highlightMenuGroups = useMemo<ContextMenuItem[][]>(() => {
    if (!highlightMenu) return [];
    const currentHl = doc?.pdfHighlights?.find((h) => h.id === highlightMenu.highlightId);
    const currentColor = currentHl?.color || "yellow";
    const cbMode = useSettingsStore.getState().colorblindMode;

    return [
      [
        {
          label: "添加 / 编辑注释",
          icon: "message-square" as any,
          onClick: () => {
            setActiveCommentHlId(highlightMenu.highlightId);
            setActiveNoteId(null);
            setHighlightMenu(null);
          },
        },
        {
          label: "移除此高亮",
          icon: "trash",
          onClick: () => {
            removePdfHighlight(docId, highlightMenu.highlightId);
            setHighlightMenu(null);
          },
        },
      ],
      [
        {
          label: cbMode ? "切换为明黄 (实线)" : "切换为黄色",
          icon: currentColor === "yellow" ? "check" : undefined,
          onClick: () => {
            updatePdfHighlight(docId, highlightMenu.highlightId, { color: "yellow" });
            setHighlightMenu(null);
          },
        },
        {
          label: cbMode ? "切换为天青蓝 (虚线)" : "切换为绿色",
          icon: currentColor === "green" ? "check" : undefined,
          onClick: () => {
            updatePdfHighlight(docId, highlightMenu.highlightId, { color: "green" });
            setHighlightMenu(null);
          },
        },
        {
          label: cbMode ? "切换为洋红紫 (双线)" : "切换为粉色",
          icon: currentColor === "pink" ? "check" : undefined,
          onClick: () => {
            updatePdfHighlight(docId, highlightMenu.highlightId, { color: "pink" });
            setHighlightMenu(null);
          },
        },
      ],
      [
        {
          label: "清除本页所有高亮",
          onClick: () => {
            clearPageHighlights(docId, highlightMenu.pageNum);
            setHighlightMenu(null);
          },
        },
        {
          label: "清除全部高亮",
          onClick: () => {
            clearAllHighlights(docId);
            setHighlightMenu(null);
          },
        },
      ],
    ];
  }, [highlightMenu, docId, doc?.pdfHighlights]);

  // 空白处/非文本区域右键菜单（添加便签附注、复制页面为图片、另存此页、顺/逆时针旋转等实用功能）
  const pageContextMenuGroups = useMemo<ContextMenuItem[][]>(() => {
    if (!pageContextMenu) return [];
    const pNum = pageContextMenu.pageNum;
    const xP = pageContextMenu.xPercent;
    const yP = pageContextMenu.yPercent;

    return [
      [
        {
          label: "在此添加便签附注",
          icon: "pin",
          onClick: () => {
            const newNote = addPdfNote(docId, pNum, xP, yP);
            setActiveNoteId(newNote.id);
            setActiveCommentHlId(null);
            setPageContextMenu(null);
          },
        },
      ],
      [
        {
          label: "复制本页为图片 (PNG)",
          icon: "copy",
          onClick: async () => {
            setPageContextMenu(null);
            const pageEl = pageRefs.current.get(pNum);
            const canvas = pageEl?.querySelector("canvas");
            if (canvas) {
              const ok = await copyCanvasToClipboard(canvas);
              if (ok) {
                await showMessage(
                  "复制成功",
                  `第 ${pNum} 页已作为高清晰度 PNG 图片复制到剪贴板，可直接在微信、文档中粘贴。`,
                );
              } else {
                await showMessage("复制失败", "无法将页面图像写入系统剪贴板。");
              }
            }
          },
        },
        {
          label: "另存此页为独立 PDF",
          icon: "download",
          onClick: () => {
            setPageContextMenu(null);
            void extractPdfPage(docId, pNum);
          },
        },
      ],
      [
        {
          label: "顺时针旋转此页 90°",
          icon: "rotate-cw",
          onClick: () => {
            setPageContextMenu(null);
            targetPageAfterReload.current = pNum;
            void rotatePdfPage(docId, pNum, true);
          },
        },
        {
          label: "逆时针旋转此页 90°",
          icon: "rotate-ccw",
          onClick: () => {
            setPageContextMenu(null);
            targetPageAfterReload.current = pNum;
            void rotatePdfPage(docId, pNum, false);
          },
        },
        {
          label: "删除当前页",
          icon: "trash",
          onClick: () => {
            setPageContextMenu(null);
            void deletePdfPage(docId, pNum);
          },
        },
      ],
      [
        {
          label: "撤销操作",
          hint: "Ctrl+Z",
          icon: "undo",
          onClick: () => {
            setPageContextMenu(null);
            targetPageAfterReload.current = currentPage;
            void undoPdf(docId);
          },
        },
        {
          label: "重做操作",
          hint: "Ctrl+Y",
          icon: "redo",
          onClick: () => {
            setPageContextMenu(null);
            targetPageAfterReload.current = currentPage;
            void redoPdf(docId);
          },
        },
      ],
      [
        {
          label: "清除本页所有高亮",
          icon: "trash",
          onClick: () => {
            setPageContextMenu(null);
            targetPageAfterReload.current = currentPage;
            void clearPageHighlights(docId, pNum);
          },
        },
        {
          label: "清除全部高亮",
          icon: "trash",
          onClick: () => {
            setPageContextMenu(null);
            targetPageAfterReload.current = currentPage;
            void clearAllHighlights(docId);
          },
        },
        {
          label: "清除历史物理高亮涂层 (恢复纯净原貌)",
          icon: "refresh",
          onClick: async () => {
            setPageContextMenu(null);
            targetPageAfterReload.current = currentPage;
            await repairBurnedPdfDocument(docId);
          },
        },
      ],
      [
        {
          label: "适合页宽",
          onClick: () => {
            setPageContextMenu(null);
            setFitMode("width");
            void updateFitWidth();
          },
        },
        {
          label: "适合整页",
          onClick: () => {
            setPageContextMenu(null);
            setFitMode("page");
            void updateFitPage();
          },
        },
      ],
    ];
  }, [pageContextMenu, docId, currentPage, updateFitWidth, updateFitPage]);

  // 根据当前选择的阅读底色，渲染舒适的背景色调
  const containerBgClass = useMemo(() => {
    if (paperTheme === "dark") return "bg-[#18181b]";
    if (paperTheme === "warm") return "bg-[#ece5d8]";
    if (paperTheme === "green") return "bg-[#ddeadf]";
    if (paperTheme === "parchment") return "bg-[#ede3cb]";
    return isDark ? "bg-[#18181b]" : "bg-neutral-100";
  }, [paperTheme, isDark]);

  return (
    <div
      onClick={() => {
        if (pane !== undefined && activePane !== pane) {
          useAppStore.getState().setActivePane(pane);
        }
      }}
      className="relative flex h-full w-full flex-col overflow-hidden bg-panel"
    >
      {/* PDF 顶置工具栏 */}
      <div className="flex h-9 shrink-0 items-center justify-between gap-1 border-b border-line bg-panel px-2 text-[12px] text-muted relative z-20">
        {/* 左侧：页码选择器 */}
        <div className="flex shrink-0 items-center gap-1">
          <button
            type="button"
            title="上一页 (Page Up)"
            disabled={currentPage <= 1}
            onClick={() => scrollToPage(currentPage - 1)}
            className="flex h-7 w-7 shrink-0 items-center justify-center rounded hover:bg-hover hover:text-fg disabled:opacity-40"
          >
            <Icon name="arrow-up" size={13} className="shrink-0" />
          </button>

          <div className="flex shrink-0 items-center gap-1">
            <input
              type="text"
              value={currentPage}
              onChange={(e) => {
                const val = parseInt(e.target.value, 10);
                if (!isNaN(val)) scrollToPage(val);
              }}
              className="h-6 w-10 shrink-0 rounded border border-line bg-app text-center text-[12px] text-fg outline-none focus:border-accent"
            />
            <span className="shrink-0 whitespace-nowrap text-muted/70 text-[11px]">/ {numPages || 1}</span>
          </div>

          <button
            type="button"
            title="下一页 (Page Down)"
            disabled={currentPage >= numPages}
            onClick={() => scrollToPage(currentPage + 1)}
            className="flex h-7 w-7 shrink-0 items-center justify-center rounded hover:bg-hover hover:text-fg disabled:opacity-40"
          >
            <Icon name="arrow-down" size={13} className="shrink-0" />
          </button>
        </div>

        {/* 中间：缩放控制（支持弹出百分比选项及自适应） & 自定义阅读底色切换器 */}
        <div className="flex shrink-0 items-center gap-1">
          <button
            type="button"
            title="缩小 (Ctrl+- 或 Ctrl+滚轮)"
            onClick={() => {
              setFitMode("custom");
              const next = Math.max(0.3, Number((scale - 0.15).toFixed(2)));
              setScale(next);
            }}
            className="flex h-7 w-7 shrink-0 items-center justify-center rounded hover:bg-hover hover:text-fg"
          >
            <Icon name="zoom-out" size={14} className="shrink-0" />
          </button>

          {/* 交互式缩放百分比菜单按钮 */}
          <button
            ref={scaleBtnRef}
            type="button"
            title="点击选择缩放比例或页面自适应"
            onClick={() => {
              if (scaleMenuPos) {
                setScaleMenuPos(null);
              } else if (scaleBtnRef.current) {
                const rect = scaleBtnRef.current.getBoundingClientRect();
                setScaleMenuPos({ left: rect.left, top: rect.bottom + 4 });
                setThemeMenuPos(null);
              }
            }}
            className="flex h-7 shrink-0 items-center gap-0.5 rounded px-1.5 font-mono text-[11.5px] text-fg/90 hover:bg-hover hover:text-accent transition-colors"
          >
            <span>{Math.round(scale * 100)}%</span>
            <Icon name="chevron-down" size={10} className="opacity-60" />
          </button>

          <button
            type="button"
            title="放大 (Ctrl+= 或 Ctrl+滚轮)"
            onClick={() => {
              setFitMode("custom");
              const next = Math.min(4.0, Number((scale + 0.15).toFixed(2)));
              setScale(next);
            }}
            className="flex h-7 w-7 shrink-0 items-center justify-center rounded hover:bg-hover hover:text-fg"
          >
            <Icon name="zoom-in" size={14} className="shrink-0" />
          </button>

          <div className="mx-1 h-4 w-px shrink-0 bg-line" />

          <button
            type="button"
            title="适合页宽"
            onClick={() => {
              setFitMode("width");
              void updateFitWidth();
            }}
            className={`flex h-7 shrink-0 whitespace-nowrap items-center rounded px-1.5 text-[11.5px] transition-colors ${
              fitMode === "width" ? "bg-accent/15 text-accent font-medium" : "hover:bg-hover hover:text-fg"
            }`}
          >
            页宽
          </button>

          <button
            type="button"
            title="适合整页"
            onClick={() => {
              setFitMode("page");
              void updateFitPage();
            }}
            className={`flex h-7 shrink-0 whitespace-nowrap items-center rounded px-1.5 text-[11.5px] transition-colors ${
              fitMode === "page" ? "bg-accent/15 text-accent font-medium" : "hover:bg-hover hover:text-fg"
            }`}
          >
            整页
          </button>

          <div className="mx-1 h-4 w-px shrink-0 bg-line" />

          {/* 自定义阅读底色切换器 */}
          <div>
            <button
              ref={themeBtnRef}
              type="button"
              title="切换阅读底色（护眼舒适）"
              onClick={() => {
                if (themeMenuPos) {
                  setThemeMenuPos(null);
                } else if (themeBtnRef.current) {
                  const rect = themeBtnRef.current.getBoundingClientRect();
                  setThemeMenuPos({ left: rect.left, top: rect.bottom + 4 });
                  setScaleMenuPos(null);
                }
              }}
              className={`flex h-7 shrink-0 items-center gap-1.5 rounded px-2 text-[11.5px] transition-colors ${
                themeMenuPos ? "bg-hover text-fg" : "hover:bg-hover hover:text-fg"
              }`}
            >
              <span
                className="h-3.5 w-3.5 rounded-full border border-line shadow-2xs shrink-0"
                style={{ backgroundColor: activePaperTheme.preview }}
              />
              <span className="whitespace-nowrap text-fg/90">{activePaperTheme.name}</span>
              <Icon name="chevron-down" size={11} className="opacity-60 shrink-0" />
            </button>
          </div>
        </div>

        {/* 右侧：编辑与页面操作（区分单页旋转与全部旋转图标） */}
        <div className="flex shrink-0 items-center gap-0.5 sm:gap-1">
          {/* 撤销 / 重做 */}
          <button
            type="button"
            title="撤销 (Ctrl+Z)"
            onClick={() => {
              targetPageAfterReload.current = currentPage;
              void undoPdf(docId);
            }}
            className="flex h-7 w-7 shrink-0 items-center justify-center rounded hover:bg-hover hover:text-fg"
          >
            <Icon name="undo" size={13} className="shrink-0" />
          </button>

          <button
            type="button"
            title="重做 (Ctrl+Y)"
            onClick={() => {
              targetPageAfterReload.current = currentPage;
              void redoPdf(docId);
            }}
            className="flex h-7 w-7 shrink-0 items-center justify-center rounded hover:bg-hover hover:text-fg"
          >
            <Icon name="redo" size={13} className="shrink-0" />
          </button>

          <div className="mx-0.5 h-4 w-px shrink-0 bg-line" />

          <button
            type="button"
            title="顺时针旋转当前页 90°"
            onClick={() => {
              targetPageAfterReload.current = currentPage;
              void rotatePdfPage(docId, currentPage, true);
            }}
            className="flex h-7 shrink-0 whitespace-nowrap items-center gap-1 rounded px-1.5 hover:bg-hover hover:text-fg"
          >
            <Icon name="rotate-cw" size={13} className="shrink-0" />
            <span className="hidden xl:inline whitespace-nowrap">旋转当前页</span>
          </button>

          <button
            type="button"
            title="顺时针旋转全部页面 90°"
            onClick={() => {
              targetPageAfterReload.current = currentPage;
              void rotateAllPdfPages(docId, true);
            }}
            className="flex h-7 shrink-0 whitespace-nowrap items-center gap-1 rounded px-1.5 hover:bg-hover hover:text-fg"
          >
            <Icon name="rotate-all" size={13} className="shrink-0" />
            <span className="hidden 2xl:inline whitespace-nowrap">旋转全部</span>
          </button>

          <button
            type="button"
            title="删除当前页"
            onClick={() => deletePdfPage(docId, currentPage)}
            className="flex h-7 shrink-0 whitespace-nowrap items-center gap-1 rounded px-1.5 text-danger/80 hover:bg-danger/10 hover:text-danger"
          >
            <Icon name="trash" size={13} className="shrink-0" />
            <span className="hidden xl:inline whitespace-nowrap">删除此页</span>
          </button>

          <button
            type="button"
            title="将当前页另存为单独的 PDF"
            onClick={() => extractPdfPage(docId, currentPage)}
            className="flex h-7 shrink-0 whitespace-nowrap items-center gap-1 rounded px-1.5 hover:bg-hover hover:text-fg"
          >
            <Icon name="download" size={13} className="shrink-0" />
            <span className="hidden 2xl:inline whitespace-nowrap">另存此页</span>
          </button>
        </div>
      </div>

      {/* 缩放比例下拉弹窗（使用 fixed 定位，绝无遮挡） */}
      {scaleMenuPos && (
        <div
          ref={scaleDropdownRef}
          style={{ left: `${scaleMenuPos.left}px`, top: `${scaleMenuPos.top}px` }}
          className="fixed z-[300] min-w-[136px] rounded-lg border border-line bg-elevated p-1 shadow-2xl backdrop-blur-md"
        >
          <div className="px-2 py-1 text-[10.5px] font-medium text-faint">缩放比例</div>
          {SCALE_PRESETS.map((opt) => {
            const isSel = fitMode === "custom" && Math.abs(scale - opt.value) < 0.04;
            return (
              <button
                key={opt.value}
                type="button"
                onClick={() => {
                  setFitMode("custom");
                  setScale(opt.value);
                  setScaleMenuPos(null);
                }}
                className={`flex w-full items-center justify-between rounded px-2 py-1 text-left font-mono text-[11.5px] transition-colors ${
                  isSel ? "bg-accent/15 font-semibold text-accent" : "hover:bg-hover hover:text-fg"
                }`}
              >
                <span>{opt.label}</span>
                {isSel && <Icon name="check" size={12} className="text-accent" />}
              </button>
            );
          })}
          <div className="my-1 h-px bg-line/60" />
          <button
            type="button"
            onClick={() => {
              setFitMode("width");
              void updateFitWidth();
              setScaleMenuPos(null);
            }}
            className={`flex w-full items-center justify-between rounded px-2 py-1 text-left text-[11.5px] transition-colors ${
              fitMode === "width" ? "bg-accent/15 font-semibold text-accent" : "hover:bg-hover hover:text-fg"
            }`}
          >
            <span>适合页宽</span>
            {fitMode === "width" && <Icon name="check" size={12} className="text-accent" />}
          </button>
          <button
            type="button"
            onClick={() => {
              setFitMode("page");
              void updateFitPage();
              setScaleMenuPos(null);
            }}
            className={`flex w-full items-center justify-between rounded px-2 py-1 text-left text-[11.5px] transition-colors ${
              fitMode === "page" ? "bg-accent/15 font-semibold text-accent" : "hover:bg-hover hover:text-fg"
            }`}
          >
            <span>适合整页</span>
            {fitMode === "page" && <Icon name="check" size={12} className="text-accent" />}
          </button>
        </div>
      )}

      {/* 阅读底色切换下拉弹窗（使用 fixed 定位，绝无遮挡） */}
      {themeMenuPos && (
        <div
          ref={themeDropdownRef}
          style={{ left: `${themeMenuPos.left}px`, top: `${themeMenuPos.top}px` }}
          className="fixed z-[300] min-w-[150px] rounded-lg border border-line bg-elevated p-1 shadow-2xl backdrop-blur-md"
        >
          <div className="px-2 py-1 text-[10.5px] font-medium text-faint">选择阅读底色</div>
          {PDF_PAPER_THEMES.map((theme) => {
            const isSel = theme.id === activePaperTheme.id;
            return (
              <button
                key={theme.id}
                type="button"
                onClick={() => {
                  useAppStore.getState().patchDoc(docId, { pdfPaperTheme: theme.id });
                  setThemeMenuPos(null);
                }}
                className={`flex w-full items-center justify-between gap-2 rounded px-2 py-1.5 text-left text-[11.5px] transition-colors ${
                  isSel ? "bg-accent/15 font-medium text-accent" : "hover:bg-hover hover:text-fg"
                }`}
              >
                <div className="flex items-center gap-2">
                  <span
                    className="h-3.5 w-3.5 rounded-full border border-line shrink-0"
                    style={{ backgroundColor: theme.preview }}
                  />
                  <span>{theme.name}</span>
                </div>
                {isSel && <Icon name="check" size={12} className="text-accent" />}
              </button>
            );
          })}
        </div>
      )}

      {/* 主视口区域：页面画布渲染流 */}
      <div
        ref={containerRef}
        onScroll={handleScroll}
        onContextMenu={handleContextMenu}
        className={`flex flex-1 flex-col items-center overflow-y-auto overflow-x-auto p-6 scrollbar-thin transition-colors duration-200 ${containerBgClass}`}
      >
        {loading ? (
          <div className="flex h-full flex-col items-center justify-center gap-3 text-muted">
            <Icon name="loader" size={28} className="animate-spin text-accent" />
            <div className="text-[13px]">正在加载并解析 PDF 文档…</div>
          </div>
        ) : error ? (
          <div className="flex h-full flex-col items-center justify-center gap-3 text-center text-muted">
            <Icon name="lock" size={32} className="text-warning/80" />
            <div className="text-[14px] font-medium text-fg">{error}</div>
            <button
              type="button"
              onClick={() => void loadPdf()}
              className="mt-2 rounded-md bg-accent px-4 py-1.5 text-[12px] text-accent-contrast shadow transition-colors hover:brightness-105"
            >
              输入密码解锁
            </button>
          </div>
        ) : (
          <div
            className={`flex flex-col items-center gap-6 transition-all ${
              paperTheme === "dark" ? "invert contrast-[0.9] hue-rotate-180" : ""
            }`}
          >
            {Array.from({ length: numPages }, (_, i) => i + 1).map((pNum) => (
              <div
                key={pNum}
                ref={(el) => {
                  if (el) pageRefs.current.set(pNum, el);
                  else pageRefs.current.delete(pNum);
                }}
                data-page-number={pNum}
                style={{ backgroundColor: activePaperTheme.color }}
                className="relative rounded shadow-lg transition-shadow"
              >
                <PdfPage
                  pdfProxy={pdfProxy}
                  pageNum={pNum}
                  scale={scale}
                  docId={docId}
                  paperTheme={paperTheme}
                  highlights={(doc?.pdfHighlights ?? []).filter((h) => h.page === pNum)}
                  notes={(doc?.pdfNotes ?? []).filter((n) => n.page === pNum)}
                  activeNoteId={activeNoteId}
                  setActiveNoteId={setActiveNoteId}
                  activeCommentHlId={activeCommentHlId}
                  setActiveCommentHlId={setActiveCommentHlId}
                  onHighlightContextMenu={(x, y, hlId, p) =>
                    setHighlightMenu({ x, y, highlightId: hlId, pageNum: p })
                  }
                />
              </div>
            ))}
          </div>
        )}
      </div>

      {/* 划词选区处理菜单 */}
      {selectionMenu ? (
        <ContextMenu
          x={selectionMenu.x}
          y={selectionMenu.y}
          groups={selectionMenuGroups}
          onClose={() => setSelectionMenu(null)}
        />
      ) : null}

      {/* 高亮标注右键菜单（支持移除） */}
      {highlightMenu ? (
        <ContextMenu
          x={highlightMenu.x}
          y={highlightMenu.y}
          groups={highlightMenuGroups}
          onClose={() => setHighlightMenu(null)}
        />
      ) : null}

      {/* 空白处/非文本区域右键菜单（添加便签附注、另存本页、旋转等） */}
      {pageContextMenu ? (
        <ContextMenu
          x={pageContextMenu.x}
          y={pageContextMenu.y}
          groups={pageContextMenuGroups}
          onClose={() => setPageContextMenu(null)}
        />
      ) : null}
    </div>
  );
}

/** 单页 Canvas + TextLayer 渲染组件（带视口虚拟化渲染、一体化缩放与护眼底色） */
function PdfPage({
  pdfProxy,
  pageNum,
  scale,
  docId,
  paperTheme,
  highlights,
  notes,
  activeNoteId,
  setActiveNoteId,
  activeCommentHlId,
  setActiveCommentHlId,
  onHighlightContextMenu,
}: {
  pdfProxy: pdfjsLib.PDFDocumentProxy | null;
  pageNum: number;
  scale: number;
  docId: string;
  paperTheme: string;
  highlights: ReturnType<typeof useAppStore.getState>["docs"][0]["pdfHighlights"];
  notes: ReturnType<typeof useAppStore.getState>["docs"][0]["pdfNotes"];
  activeNoteId: string | null;
  setActiveNoteId: (id: string | null) => void;
  activeCommentHlId: string | null;
  setActiveCommentHlId: (id: string | null) => void;
  onHighlightContextMenu: (x: number, y: number, highlightId: string, pageNum: number) => void;
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const textLayerRef = useRef<HTMLDivElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [baseSize, setBaseSize] = useState<{ width: number; height: number } | null>(null);
  const [isVisible, setIsVisible] = useState(false);
  const cbMode = useSettingsStore((s) => s.colorblindMode);

  // 1. 固有尺寸（1.0 比例下）只获取一次，后续缩放时尺寸同步计算，杜绝外框与内容分离的两步缩放
  useEffect(() => {
    if (!pdfProxy) return;
    let cancel = false;
    void pdfProxy.getPage(pageNum).then((page) => {
      if (cancel) return;
      const vp = page.getViewport({ scale: 1.0 });
      setBaseSize({ width: vp.width, height: vp.height });
    });
    return () => {
      cancel = true;
    };
  }, [pdfProxy, pageNum]);

  // 同步计算当前 scale 下的像素尺寸
  const cssWidth = baseSize ? Math.round(baseSize.width * scale) : Math.round(595 * scale);
  const cssHeight = baseSize ? Math.round(baseSize.height * scale) : Math.round(842 * scale);

  // 2. 视口可见性观察器（虚拟化渲染：远离视口的页面不渲染重负载 canvas/textLayer）
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const observer = new IntersectionObserver(
      (entries) => {
        const entry = entries[0];
        setIsVisible(entry.isIntersecting);
      },
      {
        root: null,
        rootMargin: "450px 0px 450px 0px", // 提前 450px 预加载上下页面
      },
    );
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  // 3. 可见时进行高清 Canvas + TextLayer 渲染
  useEffect(() => {
    if (!pdfProxy || !isVisible) return;
    let cancel = false;
    let renderTask: any = null;

    const render = async () => {
      try {
        const page = await pdfProxy.getPage(pageNum);
        if (cancel) return;

        const dpr = window.devicePixelRatio || 1;
        const viewport = page.getViewport({ scale: scale * dpr });
        const cssViewport = page.getViewport({ scale });

        const canvas = canvasRef.current;
        if (!canvas) return;

        const context = canvas.getContext("2d");
        if (!context) return;

        canvas.width = viewport.width;
        canvas.height = viewport.height;

        // 在尺寸变更的第一帧同步绘制纯白背景，避免重绘期间露出黑底
        context.fillStyle = "#ffffff";
        context.fillRect(0, 0, canvas.width, canvas.height);

        canvas.style.width = `${cssViewport.width}px`;
        canvas.style.height = `${cssViewport.height}px`;

        renderTask = page.render({
          canvasContext: context,
          viewport,
        });
        await renderTask.promise;
        if (cancel) return;

        // 渲染透明文本选择层（保证选区与缩放高度严格一致）
        if (textLayerRef.current) {
          textLayerRef.current.innerHTML = "";
          textLayerRef.current.style.setProperty("--scale-factor", String(cssViewport.scale));

          const textContent = await page.getTextContent();
          if (cancel) return;

          pdfjsLib.renderTextLayer({
            textContentSource: textContent,
            container: textLayerRef.current,
            viewport: cssViewport,
            textDivs: [],
          });
        }
      } catch (err: any) {
        if (err?.name !== "RenderingCancelledException") {
          console.error(`渲染第 ${pageNum} 页错误`, err);
        }
      }
    };

    void render();

    return () => {
      cancel = true;
      if (renderTask) renderTask.cancel();
    };
  }, [pdfProxy, pageNum, scale, isVisible]);

  // 护眼底色主题颜色查找
  const activeTheme = useMemo(
    () => PDF_PAPER_THEMES.find((t) => t.id === paperTheme) || PDF_PAPER_THEMES[0],
    [paperTheme],
  );

  return (
    <div
      ref={containerRef}
      style={{
        width: `${cssWidth}px`,
        height: `${cssHeight}px`,
      }}
      className="relative flex items-center justify-center shadow-xs"
    >
      {isVisible ? (
        <>
          <canvas
            ref={canvasRef}
            style={{
              width: `${cssWidth}px`,
              height: `${cssHeight}px`,
            }}
            className="block select-none"
          />

          {/* 舒适护眼纸张底色遮罩（使用 multiply 模式，保持字迹高锐度黑度，纸面呈现米黄/绿豆沙自然质感） */}
          {paperTheme !== "white" && paperTheme !== "dark" ? (
            <div
              style={{
                backgroundColor: activeTheme.color,
                mixBlendMode: "multiply",
              }}
              className="absolute inset-0 pointer-events-none z-[4]"
            />
          ) : null}

          <div
            ref={textLayerRef}
            className="textLayer absolute inset-0 select-text leading-none z-[6]"
            style={{
              width: `${cssWidth}px`,
              height: `${cssHeight}px`,
              // @ts-ignore
              "--scale-factor": String(scale),
            }}
          />

          {/* 交互式高亮标注遮罩层：防重叠加深合成图层 + 点击交互热区 + 划词注释角标 */}
          {highlights && highlights.length > 0 ? (
            <div className="absolute inset-0 pointer-events-none z-10">
              {highlights.map((hl) => {
                const hasComment = Boolean(hl.comment && hl.comment.trim());
                const firstRect = hl.rects[0];
                const solidColor = cbMode
                  ? hl.color === "green"
                    ? "#38bdf8"
                    : hl.color === "pink"
                      ? "#c084fc"
                      : "#facc15"
                  : hl.color === "green"
                    ? "#22c55e"
                    : hl.color === "pink"
                      ? "#ec4899"
                      : "#f59e0b";
                const borderBottomStyle = cbMode
                  ? hl.color === "green"
                    ? "2px dashed #0284c7"
                    : hl.color === "pink"
                      ? "3px double #9333ea"
                      : "2px solid #d97706"
                  : undefined;

                return (
                  <div key={hl.id} className="contents pointer-events-auto">
                    {/* 1. 纯净视觉混合层：统一设置 opacity 与 multiply，内部各个矩形重叠绝不加深 */}
                    <div
                      style={{ mixBlendMode: "multiply", opacity: 0.45 }}
                      className="absolute inset-0 pointer-events-none"
                    >
                      {hl.rects.map((r, i) => (
                        <div
                          key={`${hl.id}-visual-${i}`}
                          style={{
                            left: `${r.xPercent}%`,
                            top: `${r.yPercent}%`,
                            width: `${r.wPercent}%`,
                            height: `${r.hPercent}%`,
                            backgroundColor: solidColor,
                            borderBottom: borderBottomStyle,
                          }}
                          className="absolute rounded-xs"
                        />
                      ))}
                    </div>

                    {/* 2. 透明交互响应热区 */}
                    {hl.rects.map((r, i) => (
                      <div
                        key={`${hl.id}-hit-${i}`}
                        data-highlight-id={hl.id}
                        style={{
                          left: `${r.xPercent}%`,
                          top: `${r.yPercent}%`,
                          width: `${r.wPercent}%`,
                          height: `${r.hPercent}%`,
                        }}
                        className="absolute cursor-pointer rounded-xs hover:bg-black/5 dark:hover:bg-white/10 transition-colors"
                        title={
                          hasComment
                            ? `注释: ${hl.comment}\n(点击查看/编辑注释，右键切换颜色或移除)`
                            : "划词高亮 (点击查看/添加注释，右键切换颜色或移除)"
                        }
                        onClick={(e) => {
                          e.stopPropagation();
                          setActiveCommentHlId(hl.id);
                          setActiveNoteId(null);
                        }}
                        onContextMenu={(e) => {
                          e.preventDefault();
                          e.stopPropagation();
                          onHighlightContextMenu(e.clientX, e.clientY, hl.id, pageNum);
                        }}
                      />
                    ))}

                    {/* 划词高亮附带的注释图钉角标 */}
                    {hasComment && firstRect ? (
                      <div
                        style={{
                          left: `${firstRect.xPercent + firstRect.wPercent}%`,
                          top: `${firstRect.yPercent}%`,
                        }}
                        className="absolute -translate-y-1/2 -translate-x-1/2 z-20 cursor-pointer pointer-events-auto"
                        title={`查看注释: ${hl.comment}`}
                        onClick={(e) => {
                          e.stopPropagation();
                          setActiveCommentHlId(hl.id);
                          setActiveNoteId(null);
                        }}
                      >
                        <span className="flex h-4 w-4 items-center justify-center rounded-full bg-amber-500 text-white text-[9px] shadow-sm hover:scale-110 transition-transform">
                          💬
                        </span>
                      </div>
                    ) : null}

                    {/* 展开的划词高亮注释卡片 */}
                    {activeCommentHlId === hl.id ? (
                      <PdfHighlightCommentCard
                        highlight={hl}
                        docId={docId}
                        isOpen={true}
                        onClose={() => setActiveCommentHlId(null)}
                      />
                    ) : null}
                  </div>
                );
              })}
            </div>
          ) : null}

          {/* 交互式便签附注图钉与卡片 */}
          {notes && notes.length > 0 ? (
            <div className="absolute inset-0 pointer-events-none z-20">
              {notes.map((note) => (
                <PdfNoteMarker
                  key={note.id}
                  note={note}
                  docId={docId}
                  isOpen={activeNoteId === note.id}
                  onOpen={() => {
                    setActiveNoteId(note.id);
                    setActiveCommentHlId(null);
                  }}
                  onClose={() => setActiveNoteId(null)}
                />
              ))}
            </div>
          ) : null}
        </>
      ) : (
        /* 视口外虚拟骨架占位 */
        <div className="flex h-full w-full items-center justify-center text-[12px] text-faint">
          <span className="font-mono opacity-40">第 {pageNum} 页</span>
        </div>
      )}
    </div>
  );
}

/** 划词高亮附带的注释卡片组件（先添加高亮，再针对高亮进行注释） */
function PdfHighlightCommentCard({
  highlight,
  docId,
  isOpen,
  onClose,
}: {
  highlight: PdfHighlight;
  docId: string;
  isOpen: boolean;
  onClose: () => void;
}) {
  const [comment, setComment] = useState(highlight.comment?.trim() || "");
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const colorblindMode = useSettingsStore((s) => s.colorblindMode);

  useEffect(() => {
    setComment(highlight.comment?.trim() || "");
  }, [highlight.comment]);

  useEffect(() => {
    if (isOpen) {
      setTimeout(() => textareaRef.current?.focus({ preventScroll: true }), 60);
    }
  }, [isOpen]);

  const handleBlur = () => {
    updatePdfHighlight(docId, highlight.id, { comment: comment.trim() });
  };

  const firstRect = highlight.rects[0] || { xPercent: 50, yPercent: 50, wPercent: 10, hPercent: 2 };
  const isRightSide = firstRect.xPercent > 50;
  const isBottomSide = firstRect.yPercent > 65;

  const cardPosClass = `${isRightSide ? "right-2" : "left-2"} ${
    isBottomSide ? "bottom-2" : "top-2"
  }`;

  return (
    <div
      style={{
        left: `${firstRect.xPercent + (isRightSide ? 0 : firstRect.wPercent)}%`,
        top: `${firstRect.yPercent}%`,
      }}
      className="absolute -translate-y-1/2 pointer-events-auto z-40"
      onClick={(e) => e.stopPropagation()}
      onContextMenu={(e) => e.stopPropagation()}
    >
      {isOpen && (
        <div
          className={`absolute ${cardPosClass} z-50 w-72 rounded-lg border border-amber-300 bg-amber-50 p-3 shadow-2xl backdrop-blur-md text-neutral-900 dark:bg-neutral-900 dark:text-neutral-100 dark:border-amber-700/60`}
          onClick={(e) => e.stopPropagation()}
        >
          <div className="flex items-center justify-between pb-1.5 border-b border-line/60">
            <div className="flex items-center gap-1.5 text-[11px] font-semibold text-fg">
              <Icon name="message-square" size={12} className="text-accent" />
              <span>划词注释</span>
            </div>
            <div className="flex items-center gap-1">
              {(["yellow", "green", "pink"] as const).map((c) => {
                const titleLabel = colorblindMode
                  ? c === "yellow"
                    ? "明黄 (实线)"
                    : c === "green"
                      ? "天青蓝 (虚线)"
                      : "洋红紫 (双线)"
                  : c === "yellow"
                    ? "黄色"
                    : c === "green"
                      ? "绿色"
                      : "粉色";
                const dotClass = colorblindMode
                  ? c === "yellow"
                    ? "bg-yellow-400"
                    : c === "green"
                      ? "bg-sky-400"
                      : "bg-purple-400"
                  : c === "yellow"
                    ? "bg-amber-400"
                    : c === "green"
                      ? "bg-emerald-400"
                      : "bg-pink-400";
                return (
                  <button
                    key={c}
                    type="button"
                    title={titleLabel}
                    onClick={() => updatePdfHighlight(docId, highlight.id, { color: c })}
                    className={`h-3.5 w-3.5 rounded-full border transition-transform ${
                      highlight.color === c ? "scale-125 border-fg ring-1 ring-fg/30" : "border-line/60 hover:scale-110"
                    } ${dotClass}`}
                  />
                );
              })}
              <div className="mx-1 h-3 w-px bg-line/60" />
              <button
                type="button"
                title="删除此高亮及注释"
                onClick={() => {
                  removePdfHighlight(docId, highlight.id);
                  onClose();
                }}
                className="rounded p-0.5 text-danger/80 hover:bg-danger/10 hover:text-danger"
              >
                <Icon name="trash" size={12} />
              </button>
              <button
                type="button"
                title="关闭"
                onClick={onClose}
                className="rounded p-0.5 text-muted hover:bg-hover hover:text-fg"
              >
                <Icon name="x" size={12} />
              </button>
            </div>
          </div>

          {highlight.text ? (
            <div className="mt-1.5 rounded bg-black/5 dark:bg-white/5 p-1.5 text-[11px] text-muted italic line-clamp-2">
              “{highlight.text}”
            </div>
          ) : null}

          <textarea
            ref={textareaRef}
            value={comment}
            onChange={(e) => setComment(e.target.value)}
            onBlur={handleBlur}
            placeholder="写下你对此段文字的批注或笔记…"
            rows={3}
            className="mt-2 w-full resize-none rounded border border-line/70 bg-input p-1.5 text-[12px] text-fg outline-none focus:border-accent"
          />

          <div className="mt-2 flex items-center justify-between text-[10.5px] text-faint">
            <span>
              {new Date(highlight.createdAt).toLocaleTimeString([], {
                hour: "2-digit",
                minute: "2-digit",
              })}
            </span>
            <button
              type="button"
              onClick={() => {
                handleBlur();
                onClose();
              }}
              className="rounded bg-accent px-2.5 py-1 text-[11px] font-semibold text-white shadow-xs hover:brightness-110 active:brightness-95 transition-all"
            >
              完成
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

/** 交互式便签附注图钉与卡片组件（智能方位避让，杜绝遮挡截断） */
function PdfNoteMarker({
  note,
  docId,
  isOpen,
  onOpen,
  onClose,
}: {
  note: PdfNote;
  docId: string;
  isOpen: boolean;
  onOpen: () => void;
  onClose: () => void;
}) {
  const [content, setContent] = useState(note.content);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const colorblindMode = useSettingsStore((s) => s.colorblindMode);

  useEffect(() => {
    setContent(note.content);
  }, [note.content]);

  useEffect(() => {
    if (isOpen) {
      setTimeout(() => textareaRef.current?.focus({ preventScroll: true }), 60);
    }
  }, [isOpen]);

  const colorStyles = {
    yellow: {
      marker: "bg-amber-400 text-amber-950 border-amber-500 hover:bg-amber-300",
      card: "border-amber-300 bg-amber-50 text-neutral-900 dark:bg-neutral-900 dark:text-neutral-100 dark:border-amber-700/60",
    },
    blue: {
      marker: "bg-sky-400 text-sky-950 border-sky-500 hover:bg-sky-300",
      card: "border-sky-300 bg-sky-50 text-neutral-900 dark:bg-neutral-900 dark:text-neutral-100 dark:border-sky-700/60",
    },
    green: colorblindMode
      ? {
          marker: "bg-orange-500 text-white border-orange-600 hover:bg-orange-400",
          card: "border-orange-300 bg-orange-50 text-neutral-900 dark:bg-neutral-900 dark:text-neutral-100 dark:border-orange-700/60",
        }
      : {
          marker: "bg-emerald-400 text-emerald-950 border-emerald-500 hover:bg-emerald-300",
          card: "border-emerald-300 bg-emerald-50 text-neutral-900 dark:bg-neutral-900 dark:text-neutral-100 dark:border-emerald-700/60",
        },
    purple: {
      marker: "bg-purple-400 text-purple-950 border-purple-500 hover:bg-purple-300",
      card: "border-purple-300 bg-purple-50 text-neutral-900 dark:bg-neutral-900 dark:text-neutral-100 dark:border-purple-700/60",
    },
  }[note.color || "yellow"];

  const handleBlur = () => {
    if (content !== note.content) {
      updatePdfNote(docId, note.id, { content });
    }
  };

  // 关键优化：智能方位避让判定。若图钉位于页面右侧（>50%），便签卡片向左内侧展开；若位于底部（>65%），卡片向上展开
  const isRightSide = note.xPercent > 50;
  const isBottomSide = note.yPercent > 65;

  const cardPosClass = `${isRightSide ? "right-2" : "left-2"} ${
    isBottomSide ? "bottom-2" : "top-2"
  }`;

  return (
    <div
      style={{
        left: `${note.xPercent}%`,
        top: `${note.yPercent}%`,
      }}
      className="absolute -translate-x-1/2 -translate-y-1/2 pointer-events-auto z-30"
      onClick={(e) => e.stopPropagation()}
      onContextMenu={(e) => e.stopPropagation()}
    >
      {/* 便签图钉按钮 */}
      <button
        type="button"
        title={note.content ? `便签：${note.content.slice(0, 30)}` : "点击查看/编辑便签"}
        onClick={onOpen}
        className={`flex h-6 w-6 items-center justify-center rounded-full border shadow-md transition-transform hover:scale-110 active:scale-95 ${colorStyles.marker}`}
      >
        <Icon name="pin" size={13} strokeWidth={2.2} />
      </button>

      {/* 展开的便签卡片（带避让定位、超高层级与立体阴影） */}
      {isOpen && (
        <div
          className={`absolute ${cardPosClass} z-50 w-64 rounded-lg border p-3 shadow-2xl backdrop-blur-md ${colorStyles.card}`}
          onClick={(e) => e.stopPropagation()}
        >
          <div className="flex items-center justify-between pb-1.5 border-b border-line/60">
            <div className="flex items-center gap-1.5 text-[11px] font-semibold text-fg">
              <Icon name="pin" size={12} className="text-accent" />
              <span>便签附注</span>
            </div>
            {/* 颜色切换小圆点 */}
            <div className="flex items-center gap-1">
              {(["yellow", "blue", "green", "purple"] as const).map((c) => {
                const titleLabel =
                  c === "yellow"
                    ? "明黄"
                    : c === "blue"
                      ? "天青蓝"
                      : c === "green"
                        ? colorblindMode
                          ? "朱砂橙"
                          : "翠绿"
                        : "洋红紫";
                return (
                  <button
                    key={c}
                    type="button"
                    title={titleLabel}
                    onClick={() => updatePdfNote(docId, note.id, { color: c })}
                    className={`h-3.5 w-3.5 rounded-full border transition-transform ${
                      note.color === c ? "scale-125 border-fg ring-1 ring-fg/30" : "border-line/60 hover:scale-110"
                    } ${
                      c === "yellow"
                        ? "bg-amber-400"
                        : c === "blue"
                          ? "bg-sky-400"
                          : c === "green"
                            ? colorblindMode
                              ? "bg-orange-500"
                              : "bg-emerald-400"
                            : "bg-purple-400"
                    }`}
                  />
                );
              })}
              <div className="mx-1 h-3 w-px bg-line/60" />
              <button
                type="button"
                title="删除此便签"
                onClick={() => {
                  deletePdfNote(docId, note.id);
                  onClose();
                }}
                className="rounded p-0.5 text-danger/80 hover:bg-danger/10 hover:text-danger"
              >
                <Icon name="trash" size={12} />
              </button>
              <button
                type="button"
                title="关闭"
                onClick={onClose}
                className="rounded p-0.5 text-muted hover:bg-hover hover:text-fg"
              >
                <Icon name="x" size={12} />
              </button>
            </div>
          </div>

          <textarea
            ref={textareaRef}
            value={content}
            onChange={(e) => setContent(e.target.value)}
            onBlur={handleBlur}
            placeholder="在此记录对此处的批注、摘要或思考…"
            rows={3}
            className="mt-2 w-full resize-none rounded border border-line/70 bg-input p-1.5 text-[12px] text-fg outline-none focus:border-accent"
          />

          <div className="mt-2 flex items-center justify-between text-[10.5px] text-faint">
            <span>
              {new Date(note.createdAt).toLocaleTimeString([], {
                hour: "2-digit",
                minute: "2-digit",
              })}
            </span>
            <button
              type="button"
              onClick={() => {
                handleBlur();
                onClose();
              }}
              className="rounded bg-accent px-2.5 py-1 text-[11px] font-semibold text-white shadow-xs hover:brightness-110 active:brightness-95 transition-all"
            >
              完成
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
