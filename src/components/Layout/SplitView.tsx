import { useCallback, useEffect, useRef, useState } from "react";
import { CodeMirrorEditor } from "../Editor/CodeMirrorEditor";
import { MarkdownPreview } from "../Preview/MarkdownPreview";
import { useScrollSync } from "../../hooks/useScrollSync";
import { useAppStore } from "../../stores/appStore";

interface Props {
  docId: string;
  isDark: boolean;
  html: string;
  hasMath: boolean;
  hasMermaid: boolean;
  lineCount: number;
  previewRef: React.RefObject<HTMLDivElement | null>;
  /** 手动刷新用（大文档关闭实时预览） */
  livePreview: boolean;
  onRefresh: () => void;
}

const MIN_RATIO = 0.2;
const MAX_RATIO = 0.8;

/** 分屏：左源码 / 右预览，可拖动分隔条，滚动同步 */
export function SplitView({
  docId,
  isDark,
  html,
  hasMath,
  hasMermaid,
  lineCount,
  previewRef,
  livePreview,
  onRefresh,
}: Props) {
  const editorRef = useRef<HTMLDivElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [ratio, setRatio] = useState(0.5);
  const draggingRef = useRef(false);
  const syncScroll = useAppStore((s) => s.syncScroll);

  useScrollSync(editorRef, previewRef, syncScroll, lineCount, docId);

  const onMouseMove = useCallback((event: MouseEvent) => {
    if (!draggingRef.current) return;
    const rect = containerRef.current?.getBoundingClientRect();
    if (!rect) return;
    const next = (event.clientX - rect.left) / rect.width;
    setRatio(Math.min(MAX_RATIO, Math.max(MIN_RATIO, next)));
  }, []);

  const stopDrag = useCallback(() => {
    draggingRef.current = false;
    document.body.style.cursor = "";
    document.body.style.userSelect = "";
  }, []);

  useEffect(() => {
    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", stopDrag);
    return () => {
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", stopDrag);
    };
  }, [onMouseMove, stopDrag]);

  return (
    <div ref={containerRef} className="flex h-full min-h-0">
      <div
        ref={editorRef}
        className="print-hide min-w-0 overflow-hidden"
        style={{ width: `${ratio * 100}%` }}
      >
        <CodeMirrorEditor key={docId} docId={docId} isDark={isDark} />
      </div>

      <div
        role="separator"
        aria-orientation="vertical"
        onMouseDown={() => {
          draggingRef.current = true;
          document.body.style.cursor = "col-resize";
          document.body.style.userSelect = "none";
        }}
        onDoubleClick={() => setRatio(0.5)}
        className="print-hide group relative w-px shrink-0 cursor-col-resize bg-line"
        title="拖动调整宽度，双击恢复 1:1"
      >
        <div className="absolute inset-y-0 -left-1 -right-1 z-10 group-hover:bg-accent-soft" />
      </div>

      <div className="print-plain relative min-w-0 flex-1">
        {livePreview ? (
          <MarkdownPreview
            docId={docId}
            html={html}
            hasMath={hasMath}
            hasMermaid={hasMermaid}
            isDark={isDark}
            scrollRef={previewRef}
          />
        ) : (
          <div className="flex h-full flex-col items-center justify-center gap-3 bg-app text-[12px] text-muted">
            <div>文档较大（{lineCount} 行），已关闭实时预览以保证编辑流畅。</div>
            <button
              type="button"
              onClick={onRefresh}
              className="rounded-md border border-line bg-elevated px-3 py-1.5 text-[12px] text-fg hover:bg-hover"
            >
              手动刷新预览
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
