import { useEffect, useRef } from "react";
import { HEAVY_DOC_LINES } from "../utils/constants";

/**
 * 分屏滚动同步：按滚动百分比映射，使用锁避免相互触发形成死循环。
 * 大文档（> 5000 行）降低同步频率，减少卡顿。
 */
export function useScrollSync(
  editorRef: React.RefObject<HTMLElement | null>,
  previewRef: React.RefObject<HTMLElement | null>,
  enabled: boolean,
  lineCount: number,
  /** 当前文档 ID：切换文档时编辑器会整体重建（SplitView 用 key={docId}），
   *  必须把它作为依赖，否则监听器仍绑在已卸载的 .cm-scroller 上，滚动同步静默失效 */
  docId: string,
): void {
  const lockRef = useRef(false);
  const timerRef = useRef<number | null>(null);
  const heavyRef = useRef(false);

  useEffect(() => {
    heavyRef.current = lineCount > HEAVY_DOC_LINES;
  }, [lineCount]);

  useEffect(() => {
    if (!enabled) return;
    const editor = editorRef.current;
    const preview = previewRef.current;
    if (!preview) return;
    // 编辑器根元素内的实际滚动容器
    const scroller = editor?.querySelector<HTMLElement>(".cm-scroller") ?? null;
    const editorTarget: HTMLElement | null = scroller ?? editor;
    if (!editorTarget) return;

    const ratioOf = (el: HTMLElement): number => {
      const max = el.scrollHeight - el.clientHeight;
      return max > 0 ? el.scrollTop / max : 0;
    };

    const scheduleRelease = () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current);
      timerRef.current = window.setTimeout(() => {
        lockRef.current = false;
        timerRef.current = null;
      }, heavyRef.current ? 220 : 80);
    };

    const onEditorScroll = () => {
      if (lockRef.current) return;
      lockRef.current = true;
      const ratio = ratioOf(editorTarget);
      preview.scrollTop = ratio * (preview.scrollHeight - preview.clientHeight);
      scheduleRelease();
    };

    const onPreviewScroll = () => {
      if (lockRef.current) return;
      lockRef.current = true;
      const ratio = ratioOf(preview);
      editorTarget.scrollTop =
        ratio * (editorTarget.scrollHeight - editorTarget.clientHeight);
      scheduleRelease();
    };

    editorTarget.addEventListener("scroll", onEditorScroll, { passive: true });
    preview.addEventListener("scroll", onPreviewScroll, { passive: true });
    return () => {
      editorTarget.removeEventListener("scroll", onEditorScroll);
      preview.removeEventListener("scroll", onPreviewScroll);
      if (timerRef.current !== null) window.clearTimeout(timerRef.current);
      timerRef.current = null;
      lockRef.current = false;
    };
  }, [editorRef, previewRef, enabled, lineCount, docId]);
}
