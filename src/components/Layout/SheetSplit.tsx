import { useCallback, useEffect, useRef, useState } from "react";
import { CodeMirrorEditor } from "../Editor/CodeMirrorEditor";
import { SheetView } from "../Sheet/SheetView";

const MIN_RATIO = 0.2;
const MAX_RATIO = 0.8;
const DEFAULT_RATIO = 0.55;

/**
 * 表格 + 源码并排（CSV / TSV 专用）。
 * 与 Markdown 的 SplitView 不同：这里不做滚动同步（表格与文本行没有一一对应关系），
 * 两侧各自独立滚动。
 */
export function SheetSplit({ docId, isDark }: { docId: string; isDark: boolean }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const draggingRef = useRef(false);
  const [ratio, setRatio] = useState(DEFAULT_RATIO);

  const onMouseMove = useCallback((event: MouseEvent) => {
    if (!draggingRef.current) return;
    const rect = containerRef.current?.getBoundingClientRect();
    if (!rect || rect.width <= 0) return;
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
      <div className="min-w-0 overflow-hidden" style={{ width: `${ratio * 100}%` }}>
        <SheetView docId={docId} />
      </div>

      <div
        role="separator"
        aria-orientation="vertical"
        onMouseDown={() => {
          draggingRef.current = true;
          document.body.style.cursor = "col-resize";
          document.body.style.userSelect = "none";
        }}
        onDoubleClick={() => setRatio(DEFAULT_RATIO)}
        className="print-hide group relative w-px shrink-0 cursor-col-resize bg-line"
        title="拖动调整宽度，双击恢复默认"
      >
        <div className="absolute inset-y-0 -left-1 -right-1 z-10 group-hover:bg-accent-soft" />
      </div>

      <div className="min-w-0 flex-1 overflow-hidden">
        <CodeMirrorEditor key={docId} docId={docId} isDark={isDark} />
      </div>
    </div>
  );
}
