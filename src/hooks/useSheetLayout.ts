import { useCallback, useEffect, useRef } from "react";
import { useAppStore } from "../stores/appStore";
import { amendLastLayoutStep, pushLayoutStep } from "../utils/sheetHistory";

/**
 * 表格的列宽 / 行高状态（存在文档状态里，因此能进撤销栈）。
 *
 * 关键点：**一次拖动只产生一个撤销步**。
 * 网格在拖动过程中每帧都会回调（这样界面才跟手），如果每帧都压一步，拖一下就会多出几十步撤销
 * （用户实测「调整行高记录了 80 多个操作」）。这里用一个"拖动会话"把同一次拖动的回调
 * 合并进同一个撤销步：第一次压步，后续帧只更新这一步的 after，window 上松手后会话结束。
 */
export function useSheetLayout(docId: string, onHistoryChange: () => void) {
  const columnWidths = useAppStore((s) => s.docs.find((d) => d.id === docId)?.sheetColumnWidths);
  const rowHeights = useAppStore((s) => s.docs.find((d) => d.id === docId)?.sheetRowHeights);
  /** 当前正在拖动的目标（`col:3` / `row:5`）；松手后置空 */
  const sessionRef = useRef<string | null>(null);

  useEffect(() => {
    const endSession = () => {
      sessionRef.current = null;
    };
    window.addEventListener("mouseup", endSession);
    window.addEventListener("pointerup", endSession);
    window.addEventListener("pointercancel", endSession);
    return () => {
      window.removeEventListener("mouseup", endSession);
      window.removeEventListener("pointerup", endSession);
      window.removeEventListener("pointercancel", endSession);
    };
  }, []);

  const record = useCallback(
    (
      key: string,
      before: number | null,
      after: number,
      field: "sheetColumnWidths" | "sheetRowHeights",
      current: Record<number, number> | undefined,
    ) => {
      if (before === after) return;
      const index = Number(key.split(":")[1]);
      if (sessionRef.current === key && amendLastLayoutStep(docId, [{ key, after }])) {
        // 同一次拖动：只更新上一步的 after
        onHistoryChange();
      } else {
        pushLayoutStep(docId, [{ key, before, after }]);
        onHistoryChange();
      }
      sessionRef.current = key;
      useAppStore.getState().patchDoc(docId, {
        [field]: { ...(current ?? {}), [index]: after },
      });
    },
    [docId, onHistoryChange],
  );

  const handleColumnResize = useCallback(
    (index: number, width: number) => {
      record(`col:${index}`, columnWidths?.[index] ?? null, width, "sheetColumnWidths", columnWidths);
    },
    [columnWidths, record],
  );

  const handleRowResize = useCallback(
    (index: number, height: number) => {
      record(`row:${index}`, rowHeights?.[index] ?? null, height, "sheetRowHeights", rowHeights);
    },
    [record, rowHeights],
  );

  /**
   * 一键「自动调整行高」产生的一批变化：**记成一步撤销**（否则调整几十行就是几十步）。
   * 网格侧会在一次令牌处理里把所有变化一次性回调过来。
   */
  const handleRowsResize = useCallback(
    (changes: Array<{ index: number; height: number }>) => {
      if (changes.length === 0) return;
      const before = rowHeights ?? {};
      const step = changes
        .filter((change) => (before[change.index] ?? null) !== change.height)
        .map((change) => ({
          key: `row:${change.index}`,
          before: before[change.index] ?? null,
          after: change.height,
        }));
      if (step.length === 0) return;
      pushLayoutStep(docId, step);
      onHistoryChange();
      const next = { ...before };
      for (const change of changes) next[change.index] = change.height;
      useAppStore.getState().patchDoc(docId, { sheetRowHeights: next });
    },
    [docId, onHistoryChange, rowHeights],
  );

  return { columnWidths, rowHeights, handleColumnResize, handleRowResize, handleRowsResize };
}
