import { useEffect, useMemo, useRef, useState } from "react";
import { Icon, type IconName } from "../common/Icon";
import { useAppStore } from "../../stores/appStore";
import { useSettingsStore } from "../../stores/settingsStore";
import { scrollToLine } from "../../utils/editorCommands";
import { isDocumentDoc, isMarkdownDoc, isPdfDoc } from "../../utils/filePath";
import { analyzeSymbols, navSupported, type NavItem, type NavKind } from "../../utils/outline";
import type { HeadingItem } from "../../types";
import { jumpToDocxBlock, useDocxOutline } from "../Docx/docxService";
import {
  batchDeletePdfAnnotations,
  deletePdfNote,
  deletePdfPage,
  focusPdfAnnotation,
  jumpToPdfPage,
  PdfThumbnail,
  removePdfHighlight,
  reorderPdfPages,
  rotatePdfPage,
  usePdfOutline,
  usePdfProxy,
  type OutlineItem,
} from "../PDF/pdfService";
import { askConfirm, showMessage } from "../../stores/dialogStore";
import { fileName } from "../../utils/filePath";

interface Props {
  previewRef?: React.RefObject<HTMLDivElement | null>;
  standalone?: boolean;
}

type NavMode = "auto" | "heading" | "symbols";

interface DisplayItem extends NavItem {
  hasChildren: boolean;
}

const KIND_ICONS: Record<NavKind, IconName> = {
  heading: "file-text",
  class: "box",
  function: "code",
  method: "code",
  key: "key",
  table: "columns",
  selector: "hash",
  section: "list",
};

export function OutlineSidebar({ previewRef, standalone = false }: Props) {
  const doc = useAppStore((s) => s.docs.find((d) => d.id === s.activeId) ?? null);

  // 如果当前是 PDF 文档，渲染专用的 PDF 大纲书签与页面缩略图侧栏
  if (doc && isPdfDoc(doc)) {
    return <PdfOutlineSection doc={doc} standalone={standalone} />;
  }

  // 如果当前是 Word 文档，渲染专用的 Word 结构大纲侧栏
  if (doc && isDocumentDoc(doc)) {
    return <DocxOutlineSection doc={doc} standalone={standalone} />;
  }

  return <MarkdownOutlineSection doc={doc} previewRef={previewRef} standalone={standalone} />;
}

interface FlatPdfOutlineRow {
  key: string;
  title: string;
  pageIndex?: number;
  depth: number;
  hasChildren: boolean;
  collapsed: boolean;
}

/** PDF 专用大纲、缩略图与注释便签导航组件（支持完整多级树形展开与折叠） */
function PdfOutlineSection({
  doc,
  standalone,
}: {
  doc: NonNullable<ReturnType<typeof useAppStore.getState>["docs"][0]>;
  standalone?: boolean;
}) {
  const pdfProxy = usePdfProxy(doc.id);
  const pdfOutline = usePdfOutline(doc.id);
  const [tab, setTab] = useState<"outline" | "thumbnails" | "annotations">("outline");
  const [filterText, setFilterText] = useState("");
  const [collapsedKeys, setCollapsedKeys] = useState<Set<string>>(() => new Set());
  const [draggedPageIndex, setDraggedPageIndex] = useState<number | null>(null);
  const listRef = useRef<HTMLDivElement>(null);

  const numPages = doc.pdfTotalPages || pdfProxy?.numPages || 1;
  const currentPage = doc.pdfCurrentPage || 1;

  // 统计标注与便签总数
  const totalAnnotations = (doc.pdfHighlights?.length || 0) + (doc.pdfNotes?.length || 0);

  // 收集所有拥有子节点的父节点 key（用于一键全部展开/全部折叠）
  const allParentKeys = useMemo(() => {
    const keys: string[] = [];
    const walk = (nodes: OutlineItem[], prefix: string) => {
      nodes.forEach((node, idx) => {
        const key = prefix ? `${prefix}-${idx}` : `${idx}`;
        if (node.items && node.items.length > 0) {
          keys.push(key);
          walk(node.items, key);
        }
      });
    };
    walk(pdfOutline, "");
    return keys;
  }, [pdfOutline]);

  // 递归过滤并展平多级大纲树
  const flatOutlineRows = useMemo<FlatPdfOutlineRow[]>(() => {
    const q = filterText.trim().toLowerCase();

    // 搜索模式下：保留自身匹配或子孙匹配的节点，并自动展开所有匹配分支
    const filterTree = (nodes: OutlineItem[], prefix: string): Array<{ node: OutlineItem; key: string; children?: any[] }> => {
      const res: Array<{ node: OutlineItem; key: string; children?: any[] }> = [];
      nodes.forEach((node, idx) => {
        const key = prefix ? `${prefix}-${idx}` : `${idx}`;
        const childMatches = node.items && node.items.length > 0 ? filterTree(node.items, key) : [];
        const selfMatch = !q || node.title.toLowerCase().includes(q);
        if (selfMatch || childMatches.length > 0) {
          res.push({
            node,
            key,
            children: childMatches.length > 0 ? childMatches : undefined,
          });
        }
      });
      return res;
    };

    const filteredTree = filterTree(pdfOutline, "");
    const rows: FlatPdfOutlineRow[] = [];

    const flatten = (
      items: Array<{ node: OutlineItem; key: string; children?: any[] }>,
      depth: number,
    ) => {
      for (const entry of items) {
        const hasChildren = Boolean(
          (q ? entry.children : entry.node.items) &&
            ((q ? entry.children?.length : entry.node.items?.length) ?? 0) > 0,
        );
        const isCollapsed = q ? false : collapsedKeys.has(entry.key);
        rows.push({
          key: entry.key,
          title: entry.node.title,
          pageIndex: entry.node.pageIndex,
          depth,
          hasChildren,
          collapsed: isCollapsed,
        });
        if (hasChildren && !isCollapsed && entry.children) {
          flatten(entry.children, depth + 1);
        }
      }
    };

    flatten(filteredTree, 0);
    return rows;
  }, [pdfOutline, filterText, collapsedKeys]);

  // 计算当前阅读页对应的最精确书签行 key
  const activeOutlineKey = useMemo(() => {
    let bestKey: string | null = null;
    let bestPage = -1;
    for (const row of flatOutlineRows) {
      if (row.pageIndex !== undefined && row.pageIndex <= currentPage && row.pageIndex >= bestPage) {
        bestPage = row.pageIndex;
        bestKey = row.key;
      }
    }
    return bestKey;
  }, [flatOutlineRows, currentPage]);

  const toggleCollapseKey = (key: string) => {
    setCollapsedKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const allCollapsed = allParentKeys.length > 0 && collapsedKeys.size >= allParentKeys.length;
  const toggleAllCollapse = () => {
    if (allCollapsed) {
      setCollapsedKeys(new Set());
    } else {
      setCollapsedKeys(new Set(allParentKeys));
    }
  };

  // 缩略图自动滚动跟随当前高亮页
  useEffect(() => {
    if (tab === "thumbnails" && listRef.current) {
      const el = listRef.current.querySelector<HTMLElement>(`[data-page="${currentPage}"]`);
      el?.scrollIntoView({ block: "nearest", behavior: "auto" });
    }
  }, [tab, currentPage]);

  const containerClass = standalone
    ? "print-hide flex w-[280px] shrink-0 flex-col border-r border-line bg-sidebar"
    : "flex h-full min-h-0 flex-col overflow-hidden bg-sidebar";

  return (
    <div className={containerClass}>
      {/* 顶部工具栏 */}
      <div className="flex h-8 shrink-0 items-center justify-between border-b border-line px-2 text-[11px] font-medium text-faint">
        <div className="flex items-center gap-1.5">
          <span className="font-medium uppercase tracking-wide text-fg/80">
            PDF
          </span>
          <span className="rounded bg-hover px-1 py-0.5 text-[10px] text-faint">
            {numPages} 页
          </span>
        </div>

        <div className="flex items-center gap-1">
          {/* 大纲 / 缩略图 / 注释 切换 */}
          <div className="flex rounded border border-line bg-input p-0.5">
            <button
              type="button"
              onClick={() => setTab("outline")}
              title="书签大纲"
              className={`flex items-center gap-1 rounded px-1.5 py-0.5 text-[10.5px] transition-colors ${
                tab === "outline"
                  ? "bg-panel font-medium text-accent shadow-xs"
                  : "text-muted hover:text-fg"
              }`}
            >
              <Icon name="book-open" size={11} />
              <span>大纲</span>
            </button>
            <button
              type="button"
              onClick={() => setTab("thumbnails")}
              title="页面缩略图"
              className={`flex items-center gap-1 rounded px-1.5 py-0.5 text-[10.5px] transition-colors ${
                tab === "thumbnails"
                  ? "bg-panel font-medium text-accent shadow-xs"
                  : "text-muted hover:text-fg"
              }`}
            >
              <Icon name="grid" size={11} />
              <span>缩略图</span>
            </button>
            <button
              type="button"
              onClick={() => setTab("annotations")}
              title="注释与便签"
              className={`flex items-center gap-1 rounded px-1.5 py-0.5 text-[10.5px] transition-colors ${
                tab === "annotations"
                  ? "bg-panel font-medium text-accent shadow-xs"
                  : "text-muted hover:text-fg"
              }`}
            >
              <Icon name="message-square" size={11} />
              <span>注释</span>
              {totalAnnotations > 0 ? (
                <span className="rounded-full bg-accent/20 px-1 text-[9px] font-semibold text-accent leading-none">
                  {totalAnnotations}
                </span>
              ) : null}
            </button>
          </div>

          {standalone ? (
            <button
              type="button"
              title="隐藏侧栏"
              onClick={() => useAppStore.getState().setOutlineVisible(false)}
              className="rounded p-0.5 hover:bg-hover hover:text-fg"
            >
              <Icon name="x" size={12} />
            </button>
          ) : null}
        </div>
      </div>

      {tab === "outline" ? (
        <>
          {/* 搜索/过滤输入框与一键展开/折叠 */}
          <div className="flex items-center gap-1 border-b border-line/60 px-2 py-1">
            <div className="flex min-w-0 flex-1 items-center gap-1 rounded border border-line/80 bg-input px-1.5 py-0.5 text-[11px]">
              <Icon name="search" size={11} className="text-faint" />
              <input
                type="text"
                value={filterText}
                onChange={(e) => setFilterText(e.target.value)}
                placeholder="搜索书签大纲…"
                className="min-w-0 flex-1 bg-transparent text-fg outline-none placeholder:text-faint"
              />
              {filterText ? (
                <button
                  type="button"
                  onClick={() => setFilterText("")}
                  className="text-faint hover:text-fg"
                >
                  <Icon name="x" size={10} />
                </button>
              ) : null}
            </div>
            {allParentKeys.length > 0 ? (
              <button
                type="button"
                title={allCollapsed ? "展开全部书签层级" : "折叠全部子书签"}
                onClick={toggleAllCollapse}
                className="flex h-6 w-6 shrink-0 items-center justify-center rounded border border-line/70 bg-input text-muted hover:bg-hover hover:text-fg"
              >
                <Icon name={allCollapsed ? "chevron-down" : "chevron-up"} size={11} />
              </button>
            ) : null}
          </div>

          {/* 多级大纲书签树列表 */}
          <div ref={listRef} className="min-h-0 flex-1 overflow-auto p-1 font-sans">
            {flatOutlineRows.length === 0 ? (
              <div className="px-3 py-6 text-center text-[12px] text-faint">
                {filterText ? "没有匹配的书签。" : "当前 PDF 未包含书签大纲。"}
              </div>
            ) : (
              <div className="flex flex-col gap-0.5">
                {flatOutlineRows.map((row) => {
                  const active = row.key === activeOutlineKey;
                  return (
                    <div
                      key={row.key}
                      style={{ paddingLeft: `${4 + row.depth * 14}px` }}
                      onClick={() => {
                        if (row.pageIndex) jumpToPdfPage(doc.id, row.pageIndex);
                      }}
                      className={`group flex cursor-pointer items-center justify-between gap-1 rounded py-1 pr-2 text-[12px] transition-colors ${
                        active
                          ? "bg-accent-soft-strong font-medium text-accent"
                          : "text-muted hover:bg-hover hover:text-fg"
                      }`}
                    >
                      <div className="flex min-w-0 items-center gap-1">
                        {row.hasChildren ? (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              toggleCollapseKey(row.key);
                            }}
                            title={row.collapsed ? "展开子章节" : "折叠子章节"}
                            className="flex h-4 w-4 shrink-0 items-center justify-center rounded text-faint hover:bg-line/50 hover:text-fg"
                          >
                            <Icon
                              name={row.collapsed ? "chevron-right" : "chevron-down"}
                              size={11}
                            />
                          </button>
                        ) : (
                          <span className="inline-block w-4 shrink-0" />
                        )}
                        <Icon
                          name={row.depth === 0 && row.hasChildren ? "book-open" : "file-text"}
                          size={12}
                          className={`shrink-0 ${
                            active ? "text-accent" : "text-faint group-hover:text-muted"
                          }`}
                        />
                        <span
                          className={`truncate ${row.depth === 0 ? "font-medium" : ""}`}
                          title={row.title}
                        >
                          {row.title}
                        </span>
                      </div>
                      {row.pageIndex ? (
                        <span className="shrink-0 font-mono text-[10.5px] text-faint">
                          P.{row.pageIndex}
                        </span>
                      ) : null}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </>
      ) : tab === "thumbnails" ? (
        /* 页面缩略图列表 */
        <div ref={listRef} className="min-h-0 flex-1 overflow-y-auto p-2 scrollbar-thin">
          <div className="flex flex-col gap-3">
            {Array.from({ length: numPages }, (_, i) => i + 1).map((pNum) => (
              <div
                key={pNum}
                data-page={pNum}
                draggable
                onDragStart={() => setDraggedPageIndex(pNum - 1)}
                onDragOver={(e) => e.preventDefault()}
                onDrop={() => {
                  if (draggedPageIndex !== null) {
                    void reorderPdfPages(doc.id, draggedPageIndex, pNum - 1);
                    setDraggedPageIndex(null);
                  }
                }}
                onClick={() => jumpToPdfPage(doc.id, pNum)}
                className={`group relative flex cursor-pointer flex-col items-center rounded-lg border p-1.5 transition-all ${
                  currentPage === pNum
                    ? "border-accent bg-accent/10 shadow-xs"
                    : "border-line bg-panel hover:border-line-strong hover:bg-hover"
                }`}
              >
                <PdfThumbnail
                  pdfProxy={pdfProxy}
                  pageNum={pNum}
                  active={currentPage === pNum}
                />

                <div className="mt-1 flex w-full items-center justify-between px-1 text-[11px] text-muted">
                  <span className="font-mono">第 {pNum} 页</span>
                  <div className="flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
                    <button
                      type="button"
                      title="旋转此页"
                      onClick={(e) => {
                        e.stopPropagation();
                        void rotatePdfPage(doc.id, pNum, true);
                      }}
                      className="rounded p-0.5 hover:bg-accent/20 hover:text-accent"
                    >
                      <Icon name="rotate-cw" size={11} />
                    </button>
                    <button
                      type="button"
                      title="删除此页"
                      onClick={(e) => {
                        e.stopPropagation();
                        void deletePdfPage(doc.id, pNum);
                      }}
                      className="rounded p-0.5 hover:bg-danger/20 hover:text-danger"
                    >
                      <Icon name="trash" size={11} />
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      ) : (
        /* 注释与便签聚合管理面板 */
        <PdfAnnotationsView doc={doc} />
      )}
    </div>
  );
}

/** 统一标注数据项接口 */
interface AnnotationUnifiedItem {
  id: string;
  type: "highlight" | "note";
  page: number;
  yPercent: number;
  quoteText?: string;
  content: string;
  color?: string;
  createdAt?: number;
}

/** PDF 注释与便签专属管理面板组件 */
function PdfAnnotationsView({
  doc,
}: {
  doc: NonNullable<ReturnType<typeof useAppStore.getState>["docs"][0]>;
}) {
  const [filterType, setFilterType] = useState<"all" | "highlight" | "note">("all");
  const [searchText, setSearchText] = useState("");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const colorblindMode = useSettingsStore((s) => s.colorblindMode);

  // 合并高亮与便签数据
  const allItems = useMemo<AnnotationUnifiedItem[]>(() => {
    const list: AnnotationUnifiedItem[] = [];

    for (const hl of doc.pdfHighlights ?? []) {
      list.push({
        id: hl.id,
        type: "highlight",
        page: hl.page,
        yPercent: hl.rects[0]?.yPercent ?? 0,
        quoteText: hl.text,
        content: hl.comment?.trim() ?? "",
        color: hl.color,
        createdAt: hl.createdAt,
      });
    }

    for (const note of doc.pdfNotes ?? []) {
      list.push({
        id: note.id,
        type: "note",
        page: note.page,
        yPercent: note.yPercent,
        content: note.content.trim(),
        color: note.color,
        createdAt: note.createdAt,
      });
    }

    list.sort((a, b) => {
      if (a.page !== b.page) return a.page - b.page;
      return a.yPercent - b.yPercent;
    });

    return list;
  }, [doc.pdfHighlights, doc.pdfNotes]);

  // 过滤后的列表
  const filteredItems = useMemo(() => {
    const q = searchText.trim().toLowerCase();
    return allItems.filter((item) => {
      if (filterType !== "all" && item.type !== filterType) return false;
      if (!q) return true;
      const matchQuote = item.quoteText?.toLowerCase().includes(q);
      const matchContent = item.content.toLowerCase().includes(q);
      return Boolean(matchQuote || matchContent);
    });
  }, [allItems, searchText, filterType]);

  // 全选 / 反选
  const allFilteredSelected =
    filteredItems.length > 0 && filteredItems.every((it) => selectedIds.has(it.id));

  const toggleSelectAll = () => {
    if (allFilteredSelected) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredItems.map((it) => it.id)));
    }
  };

  const toggleSelectItem = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const next = new Set(selectedIds);
    if (next.has(id)) {
      next.delete(id);
    } else {
      next.add(id);
    }
    setSelectedIds(next);
  };

  // 单个删除
  const handleDeleteOne = (e: React.MouseEvent, item: AnnotationUnifiedItem) => {
    e.stopPropagation();
    if (item.type === "highlight") {
      removePdfHighlight(doc.id, item.id);
    } else {
      deletePdfNote(doc.id, item.id);
    }
    if (selectedIds.has(item.id)) {
      const next = new Set(selectedIds);
      next.delete(item.id);
      setSelectedIds(next);
    }
  };

  // 批量删除
  const handleBatchDelete = async () => {
    if (selectedIds.size === 0) return;
    const count = selectedIds.size;
    const confirmed = await askConfirm({
      title: "批量删除标注",
      message: `确定要删除已选中的 ${count} 个标注/便签吗？此操作不可撤销。`,
      confirmText: "确定删除",
      danger: true,
    });
    if (!confirmed) return;

    const highlightIds: string[] = [];
    const noteIds: string[] = [];
    for (const id of selectedIds) {
      if (doc.pdfHighlights?.some((h) => h.id === id)) {
        highlightIds.push(id);
      } else if (doc.pdfNotes?.some((n) => n.id === id)) {
        noteIds.push(id);
      }
    }
    batchDeletePdfAnnotations(doc.id, { highlightIds, noteIds });
    setSelectedIds(new Set());
  };

  // 导出为 Markdown
  const handleExportMarkdown = async () => {
    if (allItems.length === 0) return;
    const targets = selectedIds.size > 0
      ? allItems.filter((it) => selectedIds.has(it.id))
      : allItems;

    const docTitle = doc.filePath ? fileName(doc.filePath) : "PDF 文档";
    let md = `# 《${docTitle}》注释与便签笔记\n\n`;
    md += `> 共导出 ${targets.length} 条标注记录\n\n---\n\n`;

    const pageMap = new Map<number, AnnotationUnifiedItem[]>();
    for (const it of targets) {
      const arr = pageMap.get(it.page) || [];
      arr.push(it);
      pageMap.set(it.page, arr);
    }

    for (const [page, list] of pageMap) {
      md += `### 第 ${page} 页\n\n`;
      for (const it of list) {
        if (it.type === "highlight") {
          if (it.quoteText) {
            md += `> 📌 划词引用: "${it.quoteText.replace(/\n+/g, " ")}"\n\n`;
          }
          if (it.content) {
            md += `✍️ **批注**：${it.content}\n\n`;
          } else {
            md += `*(划词高亮标注)*\n\n`;
          }
        } else {
          md += `🏷️ **页面便签**：${it.content}\n\n`;
        }
      }
    }

    try {
      await navigator.clipboard.writeText(md.trim());
      await showMessage("导出成功", `已将 ${targets.length} 条笔记导出为 Markdown 格式并已复制到剪贴板！`);
    } catch {
      await showMessage("导出失败", "复制到剪贴板失败，请检查系统权限。");
    }
  };

  // 计数统计
  const highlightCount = doc.pdfHighlights?.length || 0;
  const noteCount = doc.pdfNotes?.length || 0;

  return (
    <div className="flex min-h-0 flex-1 flex-col font-sans">
      {/* 搜索与过滤切换 */}
      <div className="border-b border-line/60 p-2 space-y-1.5">
        {/* 搜索框 */}
        <div className="flex items-center gap-1 rounded border border-line/80 bg-input px-1.5 py-0.5 text-[11px]">
          <Icon name="search" size={11} className="text-faint" />
          <input
            type="text"
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
            placeholder="搜索批注或引用文字…"
            className="min-w-0 flex-1 bg-transparent text-fg outline-none placeholder:text-faint"
          />
          {searchText ? (
            <button
              type="button"
              onClick={() => setSearchText("")}
              className="text-faint hover:text-fg"
            >
              <Icon name="x" size={10} />
            </button>
          ) : null}
        </div>

        {/* 类别胶囊过滤 */}
        <div className="flex items-center justify-between text-[10.5px]">
          <div className="flex items-center gap-1">
            <button
              type="button"
              onClick={() => setFilterType("all")}
              className={`rounded px-1.5 py-0.5 transition-colors ${
                filterType === "all"
                  ? "bg-accent/15 font-medium text-accent"
                  : "text-muted hover:bg-hover hover:text-fg"
              }`}
            >
              全部 ({allItems.length})
            </button>
            <button
              type="button"
              onClick={() => setFilterType("highlight")}
              className={`rounded px-1.5 py-0.5 transition-colors ${
                filterType === "highlight"
                  ? "bg-accent/15 font-medium text-accent"
                  : "text-muted hover:bg-hover hover:text-fg"
              }`}
            >
              划词 ({highlightCount})
            </button>
            <button
              type="button"
              onClick={() => setFilterType("note")}
              className={`rounded px-1.5 py-0.5 transition-colors ${
                filterType === "note"
                  ? "bg-accent/15 font-medium text-accent"
                  : "text-muted hover:bg-hover hover:text-fg"
              }`}
            >
              便签 ({noteCount})
            </button>
          </div>

          {/* 导出 Markdown 按钮 */}
          <button
            type="button"
            onClick={handleExportMarkdown}
            disabled={allItems.length === 0}
            title={
              selectedIds.size > 0
                ? `导出选中的 ${selectedIds.size} 项为 Markdown`
                : "导出全部笔记为 Markdown"
            }
            className="flex items-center gap-1 rounded px-1.5 py-0.5 text-muted hover:bg-hover hover:text-fg disabled:opacity-40 transition-colors"
          >
            <Icon name="copy" size={11} />
            <span>导出</span>
          </button>
        </div>
      </div>

      {/* 批量管理操作条 */}
      {filteredItems.length > 0 ? (
        <div className="flex items-center justify-between border-b border-line/40 bg-sidebar/80 px-2 py-1 text-[11px] text-faint">
          <label className="flex cursor-pointer items-center gap-1.5 select-none hover:text-fg">
            <input
              type="checkbox"
              checked={allFilteredSelected}
              onChange={toggleSelectAll}
              className="rounded accent-accent h-3 w-3 cursor-pointer"
            />
            <span>{selectedIds.size > 0 ? `已选 ${selectedIds.size} 项` : "全选"}</span>
          </label>

          {selectedIds.size > 0 ? (
            <button
              type="button"
              onClick={handleBatchDelete}
              className="flex items-center gap-1 rounded bg-danger/10 px-1.5 py-0.5 text-danger hover:bg-danger/20 font-medium transition-colors"
            >
              <Icon name="trash" size={10} />
              <span>删除所选 ({selectedIds.size})</span>
            </button>
          ) : null}
        </div>
      ) : null}

      {/* 注释与便签列表内容 */}
      <div className="min-h-0 flex-1 overflow-y-auto p-2 scrollbar-thin space-y-2">
        {filteredItems.length === 0 ? (
          <div className="px-3 py-8 text-center text-[12px] text-faint">
            {searchText
              ? "未找到匹配的注释或便签。"
              : "当前暂无注释或便签。\n在正文划词或空白处右键即可添加。"}
          </div>
        ) : (
          filteredItems.map((item) => {
            const isSelected = selectedIds.has(item.id);
            const isHl = item.type === "highlight";
            const dateStr = item.createdAt
              ? new Date(item.createdAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
              : "";

            return (
              <div
                key={item.id}
                onClick={() => focusPdfAnnotation(doc.id, item.page, item.id, item.type)}
                className={`group relative flex flex-col gap-1.5 rounded-lg border p-2 text-left cursor-pointer transition-all ${
                  isSelected
                    ? "border-accent bg-accent/5 shadow-xs"
                    : "border-line bg-panel hover:border-line-strong hover:bg-hover"
                }`}
              >
                {/* 顶部元信息：勾选框、类型徽章、页码与悬浮删除 */}
                <div className="flex items-center justify-between text-[10.5px]">
                  <div className="flex items-center gap-1.5">
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onClick={(e) => toggleSelectItem(item.id, e)}
                      onChange={() => {}}
                      className="rounded accent-accent h-3 w-3 cursor-pointer"
                    />
                    {isHl ? (
                      <span className="flex items-center gap-1 rounded bg-amber-500/15 px-1 py-0.2 text-[10px] font-medium text-amber-600 dark:text-amber-400">
                        <span
                          className={`h-1.5 w-1.5 rounded-full ${
                            item.color === "green"
                              ? colorblindMode
                                ? "bg-sky-400"
                                : "bg-green-500"
                              : item.color === "pink"
                                ? colorblindMode
                                  ? "bg-purple-400"
                                  : "bg-pink-500"
                                : "bg-amber-400"
                          }`}
                        />
                        {item.content ? "批注" : "高亮"}
                      </span>
                    ) : (
                      <span className="flex items-center gap-1 rounded bg-yellow-500/15 px-1 py-0.2 text-[10px] font-medium text-yellow-600 dark:text-yellow-400">
                        <Icon name="pin" size={9} />
                        便签
                      </span>
                    )}
                  </div>

                  <div className="flex items-center gap-1">
                    <span className="font-mono text-[10.5px] text-faint group-hover:text-muted">
                      第 {item.page} 页
                    </span>
                    <button
                      type="button"
                      title="删除此标注"
                      onClick={(e) => handleDeleteOne(e, item)}
                      className="rounded p-0.5 opacity-0 group-hover:opacity-100 hover:bg-danger/20 hover:text-danger text-faint transition-opacity"
                    >
                      <Icon name="trash" size={11} />
                    </button>
                  </div>
                </div>

                {/* 划词引文（如果有） */}
                {item.quoteText ? (
                  <div className="rounded-xs border-l-2 border-amber-400/80 bg-hover/40 px-1.5 py-0.5 text-[11px] text-muted italic line-clamp-2 select-none">
                    "{item.quoteText.replace(/\s+/g, " ").trim()}"
                  </div>
                ) : null}

                {/* 批注文本或便签正文 */}
                {item.content ? (
                  <div className="text-[11.5px] font-normal text-fg leading-relaxed break-words line-clamp-3">
                    {item.content}
                  </div>
                ) : isHl ? (
                  <div className="text-[10.5px] text-faint italic select-none">
                    (划词高亮，未添加批注)
                  </div>
                ) : null}

                {/* 底部时间 */}
                {dateStr ? (
                  <div className="text-right text-[9.5px] text-faint">
                    {dateStr}
                  </div>
                ) : null}
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}

/** Markdown 与源码符号大纲组件 */
function MarkdownOutlineSection({
  doc,
  previewRef,
  standalone,
}: {
  doc: ReturnType<typeof useAppStore.getState>["docs"][0] | null;
  previewRef?: React.RefObject<HTMLDivElement | null>;
  standalone?: boolean;
}) {
  const viewMode = useAppStore((s) => s.viewMode);
  const maxLevel = useSettingsStore((s) => s.outlineMaxLevel);

  const [mode, setMode] = useState<NavMode>("auto");
  const [filterText, setFilterText] = useState("");
  const [collapsed, setCollapsed] = useState<Set<string>>(() => new Set());
  const [activeId, setActiveId] = useState<string | null>(null);
  const listRef = useRef<HTMLDivElement>(null);

  const isMarkdown = isMarkdownDoc(doc);
  const hasSymbolSupport = Boolean(doc?.filePath && navSupported(doc.filePath));

  // 当前实际采用的导航类型
  const effectiveMode =
    mode === "auto"
      ? isMarkdown
        ? "heading"
        : hasSymbolSupport
          ? "symbols"
          : "heading"
      : mode;

  // 1. Markdown 标题列表转换为通用 NavItem
  const headingItems = useMemo<NavItem[]>(() => {
    if (!doc?.headings) return [];
    return doc.headings.map((h: HeadingItem) => ({
      id: h.id,
      level: h.level,
      text: h.text,
      line: h.line,
      kind: "heading",
    }));
  }, [doc?.headings]);

  // 2. 符号列表分析
  const symbolItems = useMemo<NavItem[]>(() => {
    if (!doc?.filePath || !doc?.content) return [];
    return analyzeSymbols(doc.filePath, doc.content);
  }, [doc?.filePath, doc?.content]);

  // 原始项
  const rawItems = effectiveMode === "symbols" ? symbolItems : headingItems;

  // 折叠与筛选计算
  const flatItems = useMemo<DisplayItem[]>(() => {
    const rows: DisplayItem[] = [];
    const query = filterText.trim().toLowerCase();
    let skipLevel: number | null = null;

    for (let i = 0; i < rawItems.length; i++) {
      const item = rawItems[i];
      const next = rawItems[i + 1];

      if (item.kind === "heading" && item.level > maxLevel) continue;
      if (query && !item.text.toLowerCase().includes(query)) continue;

      if (skipLevel !== null) {
        if (item.level > skipLevel) continue;
        skipLevel = null;
      }

      rows.push({
        ...item,
        hasChildren: Boolean(next && next.level > item.level),
      });

      if (collapsed.has(item.id)) skipLevel = item.level;
    }
    return rows;
  }, [rawItems, maxLevel, filterText, collapsed]);

  // 滚动时高亮当前章节（Markdown 预览）
  useEffect(() => {
    if (viewMode === "source" || effectiveMode !== "heading" || !previewRef?.current) return;
    const scroller = previewRef.current;
    if (headingItems.length === 0) return;

    let raf = 0;
    const onScroll = () => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = 0;
        const containerTop = scroller.getBoundingClientRect().top;
        let current: string | null = null;
        for (const heading of headingItems) {
          const el = scroller.querySelector<HTMLElement>(`[id="${CSS.escape(heading.id)}"]`);
          if (!el) continue;
          if (el.getBoundingClientRect().top - containerTop <= 80) current = heading.id;
          else break;
        }
        setActiveId(current ?? headingItems[0]?.id ?? null);
      });
    };
    onScroll();
    scroller.addEventListener("scroll", onScroll, { passive: true });
    return () => {
      scroller.removeEventListener("scroll", onScroll);
      if (raf) cancelAnimationFrame(raf);
    };
  }, [previewRef, headingItems, viewMode, effectiveMode]);

  // 光标行所在项目高亮（源码模式）
  const cursorLine = doc?.cursorLine ?? 1;
  useEffect(() => {
    if (rawItems.length === 0) return;
    let current: string | null = null;
    for (const item of rawItems) {
      const targetLine = item.kind === "heading" ? item.line + 1 : item.line;
      if (targetLine <= cursorLine) current = item.id;
      else break;
    }
    setActiveId(current ?? rawItems[0]?.id ?? null);
  }, [cursorLine, rawItems]);

  useEffect(() => {
    if (!activeId || !listRef.current) return;
    const el = listRef.current.querySelector<HTMLElement>(`[data-id="${CSS.escape(activeId)}"]`);
    el?.scrollIntoView({ block: "nearest" });
  }, [activeId]);

  const jumpTo = (item: DisplayItem) => {
    if (item.kind === "heading" && viewMode !== "source" && previewRef?.current) {
      const el = previewRef.current.querySelector<HTMLElement>(`[id="${CSS.escape(item.id)}"]`);
      if (el) {
        previewRef.current.scrollTo({ top: el.offsetTop - 16, behavior: "smooth" });
      }
    } else {
      const line0 = item.kind === "heading" ? item.line : item.line - 1;
      scrollToLine(Math.max(0, line0));
    }
    setActiveId(item.id);
  };

  const containerClass = standalone
    ? "print-hide flex w-[260px] shrink-0 flex-col border-r border-line bg-sidebar"
    : "flex h-full min-h-0 flex-col overflow-hidden bg-sidebar";

  return (
    <div className={containerClass}>
      {/* 顶部工具栏 */}
      <div className="flex h-8 shrink-0 items-center justify-between border-b border-line px-2 text-[11px] font-medium text-faint">
        <div className="flex items-center gap-1.5">
          <span className="font-medium uppercase tracking-wide text-fg/80">
            {effectiveMode === "symbols" ? "符号导航" : "文档大纲"}
          </span>
          <span className="rounded bg-hover px-1 py-0.5 text-[10px] text-faint">
            {flatItems.length}
          </span>
        </div>

        <div className="flex items-center gap-1">
          <select
            value={mode}
            title="导航模式"
            onChange={(e) => setMode(e.target.value as NavMode)}
            className="h-5 rounded border border-line bg-input px-1 text-[10px] text-fg outline-none"
          >
            <option value="auto">自动</option>
            <option value="heading">大纲</option>
            <option value="symbols">符号</option>
          </select>

          {effectiveMode === "heading" ? (
            <select
              value={maxLevel}
              title="筛选标题等级"
              onChange={(e) =>
                useSettingsStore.getState().set("outlineMaxLevel", Number(e.target.value))
              }
              className="h-5 rounded border border-line bg-input px-1 text-[10px] text-fg outline-none"
            >
              <option value={6}>全部</option>
              <option value={1}>H1</option>
              <option value={2}>≤ H2</option>
              <option value={3}>≤ H3</option>
              <option value={4}>≤ H4</option>
            </select>
          ) : null}

          {standalone ? (
            <button
              type="button"
              title="隐藏侧栏"
              onClick={() => useAppStore.getState().setOutlineVisible(false)}
              className="rounded p-0.5 hover:bg-hover hover:text-fg"
            >
              <Icon name="x" size={12} />
            </button>
          ) : null}
        </div>
      </div>

      {/* 搜索/过滤输入框 */}
      <div className="border-b border-line/60 px-2 py-1">
        <div className="flex items-center gap-1 rounded border border-line/80 bg-input px-1.5 py-0.5 text-[11px]">
          <Icon name="search" size={11} className="text-faint" />
          <input
            type="text"
            value={filterText}
            onChange={(e) => setFilterText(e.target.value)}
            placeholder="过滤大纲与符号…"
            className="min-w-0 flex-1 bg-transparent text-fg outline-none placeholder:text-faint"
          />
          {filterText ? (
            <button
              type="button"
              onClick={() => setFilterText("")}
              className="text-faint hover:text-fg"
            >
              <Icon name="x" size={10} />
            </button>
          ) : null}
        </div>
      </div>

      {/* 列表渲染 */}
      {flatItems.length === 0 ? (
        <div className="px-3 py-6 text-center text-[12px] leading-relaxed text-faint">
          {filterText
            ? "没有匹配的项。"
            : effectiveMode === "symbols"
              ? "未在当前文件中识别出符号。"
              : "当前文档没有标题。"}
        </div>
      ) : (
        <div ref={listRef} className="min-h-0 flex-1 overflow-auto px-1 py-1 font-sans">
          {flatItems.map((item) => {
            const isCollapsed = collapsed.has(item.id);
            const active = activeId === item.id;
            const iconName = KIND_ICONS[item.kind] ?? "file-text";

            return (
              <div
                key={item.id}
                data-id={item.id}
                className={`group flex cursor-pointer items-center gap-1 rounded py-[2px] pr-1.5 text-[12px] transition-colors ${
                  active
                    ? "bg-accent-soft-strong font-medium text-accent"
                    : "text-muted hover:bg-hover hover:text-fg"
                }`}
                style={{ paddingLeft: 4 + Math.max(0, item.level - 1) * 12 }}
                onClick={() => jumpTo(item)}
                title={`${item.text} (行 ${item.kind === "heading" ? item.line + 1 : item.line})`}
              >
                {item.hasChildren ? (
                  <button
                    type="button"
                    tabIndex={-1}
                    className="flex h-3.5 w-3.5 shrink-0 items-center justify-center rounded text-faint hover:bg-active"
                    onClick={(e) => {
                      e.stopPropagation();
                      setCollapsed((prev) => {
                        const next = new Set(prev);
                        if (next.has(item.id)) next.delete(item.id);
                        else next.add(item.id);
                        return next;
                      });
                    }}
                  >
                    <Icon name={isCollapsed ? "chevron-right" : "chevron-down"} size={11} />
                  </button>
                ) : (
                  <span className="w-3.5 shrink-0" />
                )}

                <Icon
                  name={iconName}
                  size={12}
                  className={`shrink-0 ${active ? "text-accent" : "text-faint group-hover:text-muted"}`}
                />
                <span className="truncate">{item.text || "(未命名)"}</span>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/** Word (.docx) 专用文档大纲组件 */
function DocxOutlineSection({
  doc,
  standalone,
}: {
  doc: NonNullable<ReturnType<typeof useAppStore.getState>["docs"][0]>;
  standalone?: boolean;
}) {
  const { items, loading, partial } = useDocxOutline(doc.id);
  const [filterText, setFilterText] = useState("");
  const [maxLevel, setMaxLevel] = useState<number>(6);
  const [activeBlockIndex, setActiveBlockIndex] = useState<number | null>(null);

  const filteredItems = useMemo(() => {
    const q = filterText.trim().toLowerCase();
    return items.filter((item) => {
      if (item.level + 1 > maxLevel) return false;
      if (!q) return true;
      return item.text.toLowerCase().includes(q);
    });
  }, [items, filterText, maxLevel]);

  const containerClass = standalone
    ? "print-hide flex w-[260px] shrink-0 flex-col border-r border-line bg-sidebar"
    : "flex h-full min-h-0 flex-col overflow-hidden bg-sidebar";

  return (
    <div className={containerClass}>
      {/* 顶部工具栏 */}
      <div className="flex h-8 shrink-0 items-center justify-between border-b border-line px-2 text-[11px] font-medium text-faint">
        <div className="flex items-center gap-1.5">
          <span className="font-medium uppercase tracking-wide text-fg/80">
            Word 大纲
          </span>
          <span className="rounded bg-hover px-1 py-0.5 text-[10px] text-faint">
            {filteredItems.length}
          </span>
        </div>

        <div className="flex items-center gap-1">
          <select
            value={maxLevel}
            title="筛选标题等级"
            onChange={(e) => setMaxLevel(Number(e.target.value))}
            className="h-5 rounded border border-line bg-input px-1 text-[10px] text-fg outline-none"
          >
            <option value={6}>全部</option>
            <option value={1}>H1</option>
            <option value={2}>≤ H2</option>
            <option value={3}>≤ H3</option>
            <option value={4}>≤ H4</option>
          </select>

          {standalone ? (
            <button
              type="button"
              title="隐藏侧栏"
              onClick={() => useAppStore.getState().setOutlineVisible(false)}
              className="rounded p-0.5 hover:bg-hover hover:text-fg"
            >
              <Icon name="x" size={12} />
            </button>
          ) : null}
        </div>
      </div>

      {/* 搜索/过滤输入框 */}
      <div className="border-b border-line/60 px-2 py-1">
        <div className="flex items-center gap-1 rounded border border-line/80 bg-input px-1.5 py-0.5 text-[11px]">
          <Icon name="search" size={11} className="text-faint" />
          <input
            type="text"
            value={filterText}
            onChange={(e) => setFilterText(e.target.value)}
            placeholder="过滤 Word 标题大纲…"
            className="min-w-0 flex-1 bg-transparent text-fg outline-none placeholder:text-faint"
          />
          {filterText ? (
            <button
              type="button"
              onClick={() => setFilterText("")}
              className="text-faint hover:text-fg"
            >
              <Icon name="x" size={10} />
            </button>
          ) : null}
        </div>
      </div>

      {/* 大纲列表 */}
      <div className="min-h-0 flex-1 overflow-auto p-1 font-sans">
        {filteredItems.length === 0 ? (
          <div className="px-3 py-6 text-center text-[12px] text-faint">
            {loading
              ? "正在解析 Word 文档大纲…"
              : filterText
                ? "没有匹配的大纲标题。"
                : "当前 Word 文档未包含大纲标题。"}
          </div>
        ) : (
          <div className="flex flex-col gap-0.5">
            {filteredItems.map((item) => {
              const active = activeBlockIndex === item.index;
              return (
                <button
                  key={`${item.index}-${item.level}`}
                  type="button"
                  data-docx-outline-item="true"
                  data-docx-outline-index={item.index}
                  data-docx-outline-level={item.level}
                  title={`${item.text}\n（第 ${item.index + 1} 块，级别 H${item.level + 1}）`}
                  onClick={() => {
                    setActiveBlockIndex(item.index);
                    jumpToDocxBlock(doc.id, item.index);
                  }}
                  style={{ paddingLeft: `${4 + Math.min(6, item.level) * 12}px` }}
                  className={`group flex w-full cursor-pointer items-center justify-between gap-1 rounded py-1 pr-1.5 text-left text-[12px] transition-colors ${
                    active
                      ? "bg-accent-soft-strong font-medium text-accent"
                      : "text-muted hover:bg-hover hover:text-fg"
                  }`}
                >
                  <div className="flex min-w-0 items-center gap-1">
                    <span
                      className={`text-[9.5px] font-mono shrink-0 ${
                        active ? "text-accent" : "text-faint"
                      }`}
                    >
                      H{item.level + 1}
                    </span>
                    <span
                      className={`truncate ${item.level === 0 ? "font-medium text-fg" : ""}`}
                    >
                      {item.text || "（未命名标题）"}
                    </span>
                  </div>
                  <span className="shrink-0 font-mono text-[9.5px] text-faint opacity-60 group-hover:opacity-100">
                    #{item.index + 1}
                  </span>
                </button>
              );
            })}
          </div>
        )}
      </div>

      {partial ? (
        <div className="border-t border-line/60 bg-sidebar/60 px-2 py-1 text-[10.5px] text-faint">
          文档较大，大纲仅覆盖已读取部分。
        </div>
      ) : null}
    </div>
  );
}

