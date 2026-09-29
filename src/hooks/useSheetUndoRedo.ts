import { useEffect } from "react";

/**
 * 表格编辑的撤销 / 重做（Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z）。
 *
 * 为什么挂在 window 上而不是容器上：表格首次保存会弹确认框，点完按钮后焦点落在 body 上，
 * 挂在容器上的 onKeyDownCapture 就再也收不到按键（表现为「保存后 Ctrl+Z 失效」）。
 * 这里用窗口捕获监听 + 两条判据：
 *   1. 事件目标在表格容器内 —— 常规情况，双栏时只有被点击的那一栏响应；
 *   2. 焦点不在任何元素上（刚关掉弹窗）且本文档是当前激活文档 —— 补上弹窗后的场景。
 *
 * 单元格编辑框 / 公式栏 / 查找框内不拦截，交给输入框自己的撤销。
 */
export function useSheetUndoRedo(options: {
  /** 是否可编辑（只读文档不接管按键） */
  enabled: boolean;
  /** 本文档是否为当前激活文档（双栏时避免两个表格同时响应） */
  isActiveDoc: boolean;
  /** 表格视图的最外层容器 */
  containerRef: React.RefObject<HTMLElement | null>;
  /** 执行一次撤销，返回是否真的撤销了（false 时不吞掉按键） */
  undo: () => boolean;
  redo: () => boolean;
}): void {
  const { enabled, isActiveDoc, containerRef, undo, redo } = options;

  useEffect(() => {
    if (!enabled) return;

    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
      const target = event.target as HTMLElement | null;
      if (
        target &&
        (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)
      ) {
        return;
      }

      const inside = Boolean(target && containerRef.current?.contains(target));
      // 只要焦点不在输入框里就允许撤销：保存确认弹窗关闭后焦点可能落在工具条按钮或 body 上，
      // 早期版本要求「事件必须在表格容器内」，导致保存后 Ctrl+Z 失效。
      // 双栏时用「当前激活文档」区分，避免两个表格抢按键。
      if (!isActiveDoc && !inside) return;

      const key = event.key.toLowerCase();
      const isUndo = key === "z" && !event.shiftKey;
      const isRedo = key === "y" || (key === "z" && event.shiftKey);
      if (!isUndo && !isRedo) return;

      if (isUndo ? undo() : redo()) {
        event.preventDefault();
        event.stopPropagation();
      }
    };

    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [enabled, isActiveDoc, containerRef, undo, redo]);
}
