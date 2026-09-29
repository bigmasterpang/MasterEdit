import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import {
  applyFormulaSuggestion,
  matchFormulaFunctions,
  replaceOrInsertReference,
  type FormulaFunction,
  type FormulaSuggestion,
} from "../../utils/formulaFunctions";

/**
 * 公式补全（输入 `=` 后按前缀提示常用函数）。
 *
 * 只做名称与用法提示：不解析参数、不计算公式，选中后插入 `函数名(`（无参函数插入 `()`），
 * 光标停在括号内等着填参数。
 */

interface SuggestState {
  suggestion: FormulaSuggestion;
  /** 当前高亮的候选项下标 */
  index: number;
  /** 输入框的位置（面板用 fixed 定位，避免被祖先的 overflow 裁剪） */
  rect: { left: number; top: number; width: number };
}

export function useFormulaSuggest(options: {
  /** 只读 / 非编辑状态不提示 */
  enabled: boolean;
  /** 输入框当前文本（用于在内容变化后重算建议） */
  value: string;
  onChange: (value: string) => void;
  inputRef: React.RefObject<HTMLInputElement | null>;
  /** 当前工作表的已用行列数：用于生成整列/整行/已用区域引用候选 */
  context?: { usedRows?: number; usedCols?: number };
}) {
  const { enabled, value, onChange, inputRef, context } = options;
  const [state, setState] = useState<SuggestState | null>(null);
  /** 用户按 Esc 主动关掉后，在本次输入上下文里不再自动弹出（换格/继续输入会重置） */
  const dismissedRef = useRef(false);
  const usedRows = context?.usedRows ?? 0;
  const usedCols = context?.usedCols ?? 0;

  const refresh = useCallback(() => {
    const input = inputRef.current;
    if (!enabled || !input || dismissedRef.current) {
      setState(null);
      return;
    }
    const caret = input.selectionStart ?? input.value.length;
    const suggestion = matchFormulaFunctions(input.value, caret, { usedRows, usedCols });
    if (!suggestion) {
      setState(null);
      return;
    }
    const rect = input.getBoundingClientRect();
    setState((prev) => ({
      suggestion,
      index: prev && prev.suggestion.query === suggestion.query ? Math.min(prev.index, suggestion.items.length - 1) : 0,
      rect: { left: rect.left, top: rect.bottom + 2, width: rect.width },
    }));
  }, [enabled, inputRef, usedCols, usedRows]);

  /* 文本变化（含切格重新载入）后重算 */
  useEffect(() => {
    refresh();
  }, [refresh, value]);

  /* 窗口尺寸/滚动变化时跟随输入框 */
  useEffect(() => {
    if (!state) return;
    const onMove = () => refresh();
    window.addEventListener("resize", onMove);
    window.addEventListener("scroll", onMove, true);
    return () => {
      window.removeEventListener("resize", onMove);
      window.removeEventListener("scroll", onMove, true);
    };
  }, [refresh, state]);

  /** 插入候选项；不传则插入当前高亮项 */
  const accept = useCallback(
    (fn?: FormulaFunction) => {
      const input = inputRef.current;
      setState((prev) => {
        if (!prev) return null;
        const picked = fn ?? prev.suggestion.items[prev.index];
        if (!picked) return null;
        const { text, caret } = applyFormulaSuggestion(input?.value ?? value, prev.suggestion, picked);
        onChange(text);
        // 光标停在括号后：写在下一帧，等受控值渲染完
        requestAnimationFrame(() => {
          const target = inputRef.current;
          if (target) {
            target.focus();
            target.setSelectionRange(caret, caret);
          }
        });
        return null;
      });
    },
    [inputRef, onChange, value],
  );

  const close = useCallback(() => {
    dismissedRef.current = true;
    setState(null);
  }, []);

  /** 新的输入上下文（例如切换单元格）后恢复自动提示 */
  const reset = useCallback(() => {
    dismissedRef.current = false;
    setState(null);
  }, []);

  /**
   * 键盘处理：面板打开时接管 ↑ ↓ Tab Enter Esc。
   * 返回 true 表示按键已被建议列表消费，调用方不要再处理。
   */
  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLInputElement>): boolean => {
      if (!state) return false;
      const total = state.suggestion.items.length;
      if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        event.preventDefault();
        const delta = event.key === "ArrowDown" ? 1 : -1;
        setState((prev) =>
          prev ? { ...prev, index: (prev.index + delta + total) % total } : prev,
        );
        return true;
      }
      if (event.key === "Tab" || (event.key === "Enter" && !event.shiftKey && total > 0)) {
        // Enter 在提示打开时先补全；补全后光标在括号内，再按 Enter 才是提交
        event.preventDefault();
        accept();
        return true;
      }
      if (event.key === "Escape") {
        event.preventDefault();
        close();
        return true;
      }
      return false;
    },
    [accept, close, state],
  );

  /** 在光标处插入一段文本（用于「插入选区引用」），并把光标移到插入内容之后 */
  const insertAtCaret = useCallback(
    (snippet: string) => {
      const input = inputRef.current;
      if (!input) return;
      const caret = input.selectionStart ?? input.value.length;
      const next = input.value.slice(0, caret) + snippet + input.value.slice(caret);
      onChange(next);
      requestAnimationFrame(() => {
        const target = inputRef.current;
        if (target) {
          target.focus();
          const pos = caret + snippet.length;
          target.setSelectionRange(pos, pos);
        }
      });
    },
    [inputRef, onChange],
  );

  /**
   * 写入一段区域引用（Excel 的"用鼠标选范围"）：
   * 光标前已有引用就替换它，否则插入。用于网格侧拖选后回填公式。
   */
  const insertReferenceText = useCallback(
    (reference: string) => {
      const input = inputRef.current;
      if (!input) return;
      const caret = input.selectionStart ?? input.value.length;
      const { text, caret: nextCaret } = replaceOrInsertReference(input.value, caret, reference);
      onChange(text);
      requestAnimationFrame(() => {
        const target = inputRef.current;
        if (target) {
          target.focus();
          target.setSelectionRange(nextCaret, nextCaret);
        }
      });
    },
    [inputRef, onChange],
  );

  /** 内容栏当前是否处于"编辑公式"状态（供网格决定是否开启拾取模式） */
  const editingFormula = enabled && /^\s*=/.test(value);

  return {
    state,
    refresh,
    accept,
    close,
    reset,
    handleKeyDown,
    insertAtCaret,
    insertReferenceText,
    editingFormula,
  };
}

/** 候选项面板（fixed 定位，跟随输入框） */
export function FormulaSuggestPanel({
  state,
  onPick,
}: {
  state: SuggestState | null;
  onPick: (fn: FormulaFunction) => void;
}) {
  const listRef = useRef<HTMLDivElement>(null);

  /* 键盘移动高亮时把该项滚进可视区 */
  useLayoutEffect(() => {
    const active = listRef.current?.querySelector<HTMLElement>("[data-active='true']");
    active?.scrollIntoView({ block: "nearest" });
  }, [state?.index, state?.suggestion.query]);

  if (!state) return null;
  return (
    <div
      ref={listRef}
      role="listbox"
      aria-label="公式提示"
      style={{ left: state.rect.left, top: state.rect.top, minWidth: Math.max(320, state.rect.width) }}
      className="fixed z-[190] max-h-[260px] overflow-y-auto rounded-[var(--radius)] border border-line bg-elevated py-1 shadow-[var(--shadow)]"
    >
      {state.suggestion.items.map((fn, index) => (
        <button
          key={fn.name}
          type="button"
          data-active={index === state.index}
          onMouseDown={(event) => {
            // 用 mouseDown 抢在输入框失焦之前
            event.preventDefault();
            onPick(fn);
          }}
          className={`flex w-full items-baseline gap-2 px-2 py-1 text-left text-[12px] ${
            index === state.index ? "bg-accent/15 text-accent" : "text-fg hover:bg-hover"
          }`}
        >
          <span className="shrink-0 font-mono font-medium">{fn.name}</span>
          <span className="shrink-0 text-[11px] text-muted">{fn.description}</span>
          <span className="min-w-0 flex-1 truncate text-right font-mono text-[11px] text-faint">
            {fn.signature}
          </span>
        </button>
      ))}
    </div>
  );
}
