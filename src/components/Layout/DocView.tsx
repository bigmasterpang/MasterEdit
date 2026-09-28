import { useState, useRef, useEffect, lazy, Suspense } from "react";
import { TabBar } from "../Tabs/TabBar";
import { MarkdownPreview } from "../Preview/MarkdownPreview";
import { CodeMirrorEditor } from "../Editor/CodeMirrorEditor";
import { SplitView } from "./SplitView";
import { WelcomeScreen } from "../Welcome/WelcomeScreen";
import { SearchBar } from "../SearchBar/SearchBar";
import { Icon } from "../common/Icon";

const PdfViewer = lazy(() => import("../PDF/PdfViewer").then((m) => ({ default: m.PdfViewer })));
import { useMarkdown, type MarkdownResult } from "../../hooks/useMarkdown";
import { useAppStore, getDocById } from "../../stores/appStore";
import { useSettingsStore } from "../../stores/settingsStore";
import { isMarkdownDoc, isPdfDoc } from "../../utils/filePath";
import { REALTIME_PREVIEW_LIMIT } from "../../utils/constants";
import { useTabDragStore } from "../../stores/tabDragStore";

interface DocViewProps {
  docId: string | null;
  pane: 0 | 1;
  isDark: boolean;
  previewRef: React.RefObject<HTMLDivElement | null>;
}

export function DocView({ docId, pane, isDark, previewRef }: DocViewProps) {
  const doc = useAppStore((s) => s.docs.find((d) => d.id === docId) ?? null);
  const defaultFontSize = useSettingsStore((s) => s.fontSize);
  const effectiveFontSize = doc?.fontSize ?? defaultFontSize;
  const viewMode = useAppStore((s) => s.viewMode);
  const layout = useAppStore((s) => s.layout);
  const activePane = layout.activePane;
  const dragStore = useTabDragStore();

  const content = doc?.content ?? "";
  const isMarkdown = isMarkdownDoc(doc);
  const rendered = useMarkdown(isMarkdown ? content : "", viewMode);
  const lineCount = doc ? doc.content.split("\n").length : 0;
  const livePreview = content.length <= REALTIME_PREVIEW_LIMIT;

  const [snapshot, setSnapshot] = useState<MarkdownResult | null>(null);
  const effective: MarkdownResult = livePreview && !snapshot ? rendered : snapshot ?? rendered;

  // 切换文档时丢弃上一份手动刷新快照：DocView 不随 docId 重建，
  // 否则大文档刷新后切到别的文档会一直显示（并被导出）上一份文档的预览
  useEffect(() => {
    setSnapshot(null);
  }, [docId]);

  // 拖动标签放置区：拖拽本身由 TabBar 的指针事件驱动（窗口开启了 dragDropEnabled，
  // Windows 上 HTML5 drag/drop 事件不会派发，因此这里只负责渲染放置提示）
  const viewRef = useRef<HTMLDivElement>(null);

  // 分栏内文档独立滚轮缩放（仅缩放鼠标所在分栏的当前文档，不影响另一栏）
  useEffect(() => {
    const el = viewRef.current;
    if (!el || !docId) return;
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) return;
      const currentDoc = getDocById(docId);
      if (!currentDoc || isPdfDoc(currentDoc)) return;
      event.preventDefault();
      event.stopPropagation();
      const base = currentDoc.fontSize ?? useSettingsStore.getState().fontSize;
      const delta = event.deltaY < 0 ? 1 : -1;
      const nextSize = Math.min(32, Math.max(10, base + delta));
      const store = useAppStore.getState();
      if (store.layout.activePane !== pane) {
        store.setActivePane(pane);
      }
      store.patchDoc(docId, { fontSize: nextSize });
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  }, [docId, pane]);

  const refreshPreview = () => {
    setSnapshot({
      html: rendered.html,
      headings: rendered.headings,
      frontMatter: rendered.frontMatter,
      frontMatterRaw: rendered.frontMatterRaw,
      hasMath: rendered.hasMath,
      hasMermaid: rendered.hasMermaid,
    });
  };

  const previewNode = (
    <MarkdownPreview
      docId={docId}
      html={effective.html}
      hasMath={effective.hasMath}
      hasMermaid={effective.hasMermaid}
      isDark={isDark}
      scrollRef={previewRef}
    />
  );

  return (
    <div
      ref={viewRef}
      data-pane-viewport={pane}
      style={{ "--editor-size": `${effectiveFontSize}px` } as React.CSSProperties}
      onClick={() => {
        if (activePane !== pane) {
          useAppStore.getState().setActivePane(pane);
        }
      }}
      className={`relative flex h-full min-w-0 flex-1 flex-col overflow-hidden bg-app ${
        activePane === pane && layout.split ? "ring-1 ring-inset ring-accent/30" : ""
      }`}
    >
      {/* 标签栏 */}
      <TabBar pane={pane} />

      {/* 文档视口内容 */}
      <div className="relative min-h-0 flex-1 overflow-hidden">
        {!doc ? (
          pane === 0 ? (
            <WelcomeScreen />
          ) : (
            <div className="flex h-full flex-col items-center justify-center gap-3 p-8 text-center text-muted">
              <Icon name="columns" size={32} className="opacity-40" />
              <div className="text-[13px] font-medium text-fg/80">右侧分栏</div>
              <div className="max-w-[280px] text-[12px] leading-relaxed text-faint">
                拖拽标签到此区域形成双栏，或在左栏标签右键选择「移到右栏」。
              </div>
            </div>
          )
        ) : isPdfDoc(doc) ? (
          <Suspense
            fallback={
              <div className="flex h-full flex-col items-center justify-center gap-2 text-muted">
                <Icon name="loader" size={24} className="animate-spin text-accent" />
                <div className="text-[12px]">加载 PDF 模块…</div>
              </div>
            }
          >
            <PdfViewer key={doc.id} docId={doc.id} pane={pane} isDark={isDark} />
          </Suspense>
        ) : !isMarkdown ? (
          <CodeMirrorEditor key={doc.id} docId={doc.id} isDark={isDark} />
        ) : viewMode === "preview" ? (
          livePreview ? (
            previewNode
          ) : (
            <div className="flex h-full flex-col items-center justify-center gap-3 text-[12px] text-muted">
              <div>文档较大（{lineCount} 行），已关闭实时预览。</div>
              <button
                type="button"
                onClick={refreshPreview}
                className="rounded-md border border-line bg-elevated px-3 py-1.5 text-fg hover:bg-hover"
              >
                渲染预览
              </button>
            </div>
          )
        ) : viewMode === "source" ? (
          <CodeMirrorEditor key={doc.id} docId={doc.id} isDark={isDark} />
        ) : (
          <SplitView
            docId={doc.id}
            isDark={isDark}
            html={effective.html}
            hasMath={effective.hasMath}
            hasMermaid={effective.hasMermaid}
            lineCount={lineCount}
            previewRef={previewRef}
            livePreview={livePreview}
            onRefresh={refreshPreview}
          />
        )}

        {/* 查找替换浮层（仅在当前获得焦点的分栏中激活） */}
        {doc && activePane === pane ? <SearchBar /> : null}
      </div>

      {/* 拖动标签分栏提示区 */}
      {(dragStore.isDragging && dragStore.targetPane === pane && dragStore.fromPane !== pane) ? (
        <div
          className="pointer-events-none absolute bottom-0 top-0 inset-0 z-50 flex items-center justify-center border-2 border-dashed border-accent bg-accent/15 backdrop-blur-[1px] transition-all"
        >
          <div className="flex items-center gap-2 rounded-lg bg-elevated/95 px-4 py-2 text-[13px] font-semibold text-accent shadow-lg border border-accent/40">
            <Icon name="columns" size={16} />
            <span>
              {layout.split
                ? pane === 0 ? "移到左栏" : "移到右栏"
                : "移到右栏 (双栏并排)"}
            </span>
          </div>
        </div>
      ) : null}
    </div>
  );
}
