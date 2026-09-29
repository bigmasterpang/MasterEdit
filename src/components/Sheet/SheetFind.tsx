import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Icon } from "../common/Icon";

/*
 * 表格查找 / 替换（Notepad++ 风格浮动面板，语义对齐 md 文档的 SearchBar）。
 *
 * ── 分工 ────────────────────────────────────────────────────────────────
 * 本文件只管**状态机 + 面板 UI**：查找/替换的真正执行由父组件注入
 * （CSV 在 JS 里扫内存表；xlsx 走 Rust `spreadsheet_find` 与待提交编辑）。
 *
 * ── 两条硬限制（面板上会如实提示）────────────────────────────────────────
 * 1. **替换只作用于当前工作表**：xlsx 的编辑只落到「当前表」，跨工作表的命中只能查找、不能替换
 *    （替换回调由父组件跳过跨表命中；面板底部有一行说明）。
 * 2. **xlsx 不支持正则**：Rust `spreadsheet_find` 只有 matchCase / wholeCell。
 *    父组件对 xlsx 传 `supportsRegex: false`，面板会把 `.*` 置灰并给出 tooltip 说明。
 *
 * ── 父组件接线契约 ──────────────────────────────────────────────────────
 * ```tsx
 * const findApi = useSheetFind({
 *   isActiveDoc,                       // 双栏时只有激活的那个表格响应快捷键
 *   supportsRegex: !isWorkbook,        // CSV = true，xlsx = false
 *   find: async (query, options, scope) => ({ hits, capped, error? }),
 *   jump: (hit) => { ...定位到该格... },
 *   replaceCurrent: async (replacement, ctx) => 替换处数,   // ctx.hit 是当前命中
 *   replaceAll: async (replacement, ctx) => 替换处数,       // ctx.hits 是本次命中的全部
 * });
 * return <SheetFindBar api={findApi} supportsRegex={!isWorkbook} />;
 * ```
 * `find` 的 options 一定是面板上的当前开关；`scope` 是 "sheet" | "selection"。
 * 替换回调可以是同步函数（返回 number）或异步（返回 Promise<number>）；
 * 返回 0 表示这一处没有真正变化（面板会提示「内容未变化」）。
 *
 * ── 快捷键 ──────────────────────────────────────────────────────────────
 * Ctrl+F 打开、Ctrl+H 打开并展开替换行、F3 / Shift+F3 下一个 / 上一个、Esc 关闭、
 * 查找框 Enter = 查找/下一个、Shift+Enter = 上一个、替换框 Enter = 替换当前并跳到下一个。
 */

/** 一条查找命中（sheet 仅 xlsx 需要；CSV 只有行列） */
export interface SheetFindHit {
  sheet?: string;
  row: number;
  col: number;
}

/** 查找开关（面板上的四个按钮） */
export interface SheetFindOptions {
  /** 区分大小写 */
  matchCase: boolean;
  /** 全字匹配（整格内容完全相同） */
  wholeCell: boolean;
  /** 正则表达式（xlsx 不支持） */
  regex: boolean;
  /** 循环查找：到达末尾后回到开头 */
  wrap: boolean;
}

/** 查找范围：整张工作表 / 当前选区 */
export type SheetFindScope = "sheet" | "selection";

/** `find` 的返回：命中列表 + 是否被上限截断 + 可选错误信息（如后端不支持某个开关） */
export interface SheetFindResult {
  hits: SheetFindHit[];
  capped: boolean;
  error?: string;
}

/** 替换回调的上下文（父组件据此知道当前开关、范围与命中集合） */
export interface SheetFindContext {
  options: SheetFindOptions;
  scope: SheetFindScope;
  /** 本次查找的全部命中 */
  hits: SheetFindHit[];
  /** 当前命中；`replaceAll` 时为 null */
  hit: SheetFindHit | null;
}

export interface SheetFindState {
  open: boolean;
  /** 替换行是否展开（Ctrl+H 展开） */
  replaceOpen: boolean;
  query: string;
  replacement: string;
  /** 当前结果对应的查询词：与 query 不一致说明用户改了内容、还没重新查找 */
  resultQuery: string;
  options: SheetFindOptions;
  scope: SheetFindScope;
  hits: SheetFindHit[];
  /** 当前命中下标（-1 表示还没有结果） */
  index: number;
  busy: boolean;
  /** 命中数达到上限被截断 */
  capped: boolean;
  /** 正则非法等错误（有错误时禁用替换） */
  error: string;
  /** 一次性提示：已替换 N 处 / 已到末尾 / 没有可替换的命中 */
  status: string;
  /** 父组件是否注入了替换能力（false 时替换按钮禁用） */
  canReplace: boolean;
}

export interface UseSheetFindOptions {
  /** 是否为当前激活文档（双栏时只有一个表格响应快捷键） */
  isActiveDoc: boolean;
  /** 当前文档是否支持正则（CSV true / xlsx false）；false 时正则开关恒为关且面板置灰 */
  supportsRegex?: boolean;
  /** 执行查找，返回命中列表（按扫描顺序） */
  find: (
    query: string,
    options: SheetFindOptions,
    scope: SheetFindScope,
  ) => Promise<SheetFindResult> | SheetFindResult;
  /** 跳到某条命中（父组件负责切表、设置活动单元格与选区） */
  jump: (hit: SheetFindHit) => void;
  /** 替换当前命中，返回真正替换的处数（0 = 内容没变）；不注入则替换按钮禁用 */
  replaceCurrent?: (replacement: string, context: SheetFindContext) => Promise<number> | number;
  /** 替换全部命中，返回真正替换的处数；不注入则替换按钮禁用 */
  replaceAll?: (replacement: string, context: SheetFindContext) => Promise<number> | number;
}

export interface SheetFindApi {
  state: SheetFindState;
  /** 用当前开关/范围重新查找（可传临时查询词 / 开关 / 范围，供刚改完就重查的场景） */
  run: (query?: string, options?: SheetFindOptions, scope?: SheetFindScope) => Promise<void>;
  /** 下一个 / 上一个（查询词改过时会自动重新查找） */
  next: (delta: number) => void;
  open: () => void;
  close: () => void;
  setQuery: (query: string) => void;
  toggleOption: (key: keyof SheetFindOptions) => void;
  setScope: (scope: SheetFindScope) => void;
  setReplacement: (value: string) => void;
  toggleReplace: () => void;
  replaceCurrent: () => Promise<void>;
  replaceAll: () => Promise<void>;
}

export const DEFAULT_SHEET_FIND_OPTIONS: SheetFindOptions = {
  matchCase: false,
  wholeCell: false,
  regex: false,
  wrap: true,
};

/** 正则是否合法；非法时返回可读的错误文案 */
function regexError(pattern: string, options: SheetFindOptions): string {
  try {
    // eslint-disable-next-line no-new
    new RegExp(pattern, options.matchCase ? "" : "i");
    return "";
  } catch (error) {
    return `正则表达式非法：${error instanceof Error ? error.message : String(error)}`;
  }
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * 表格查找 / 替换的状态机。查找与替换的具体执行由调用方注入（见文件头契约）。
 */
export function useSheetFind(options: UseSheetFindOptions): SheetFindApi {
  const {
    isActiveDoc,
    supportsRegex = true,
    find,
    jump,
    replaceCurrent: replaceCurrentInjected,
    replaceAll: replaceAllInjected,
  } = options;
  const [state, setState] = useState<SheetFindState>({
    open: false,
    replaceOpen: false,
    query: "",
    replacement: "",
    resultQuery: "",
    // 不支持正则的文档：正则开关恒为关（父组件即使传了也别打开）
    options: { ...DEFAULT_SHEET_FIND_OPTIONS, regex: false },
    scope: "sheet",
    hits: [],
    index: -1,
    busy: false,
    capped: false,
    error: "",
    status: "",
    canReplace: Boolean(replaceCurrentInjected || replaceAllInjected),
  });

  /** 回调里读最新状态（避免把 state 塞进 useCallback 依赖导致回调不断重建） */
  const stateRef = useRef(state);
  stateRef.current = state;
  const findRef = useRef(find);
  const jumpRef = useRef(jump);
  const replaceCurrentRef = useRef(replaceCurrentInjected);
  const replaceAllRef = useRef(replaceAllInjected);
  findRef.current = find;
  jumpRef.current = jump;
  replaceCurrentRef.current = replaceCurrentInjected;
  replaceAllRef.current = replaceAllInjected;

  /**
   * 执行查找并跳到第一条。
   * 开关 / 范围刚被改动时要把**新值**显式传进来（state 的 ref 要等下一次渲染才更新，
   * 否则「切一下开关」会用旧开关重查一遍，面板看起来毫无反应）。
   */
  const run = useCallback(
    async (queryOverride?: string, optionsOverride?: SheetFindOptions, scopeOverride?: SheetFindScope) => {
      const current = stateRef.current;
      const trimmed = (queryOverride ?? current.query).trim();
      const options = optionsOverride ?? current.options;
      const scope = scopeOverride ?? current.scope;
      if (trimmed === "") {
        setState((prev) => ({
          ...prev,
          hits: [],
          index: -1,
          busy: false,
          capped: false,
          error: "",
          status: "",
          resultQuery: "",
        }));
        return;
      }
      // 正则非法：本地就拦下来，不浪费一次全表扫描，并让替换按钮禁用
      if (options.regex) {
        const invalid = regexError(trimmed, options);
        if (invalid) {
          setState((prev) => ({
            ...prev,
            hits: [],
            index: -1,
            busy: false,
            capped: false,
            error: invalid,
            status: "",
            resultQuery: trimmed,
          }));
          return;
        }
      }
      setState((prev) => ({ ...prev, busy: true, error: "", status: "" }));
      try {
        const result = await findRef.current(trimmed, options, scope);
        const hits = result.hits ?? [];
        setState((prev) => ({
          ...prev,
          hits,
          index: hits.length > 0 ? 0 : -1,
          busy: false,
          capped: result.capped,
          error: result.error ?? "",
          status: "",
          // 记下这批结果对应的查询词，供「改了内容后自动重查」判断
          resultQuery: trimmed,
        }));
        if (hits.length > 0) jumpRef.current(hits[0]);
      } catch (error) {
        setState((prev) => ({
          ...prev,
          hits: [],
          index: -1,
          busy: false,
          capped: false,
          error: messageOf(error),
          status: "",
          resultQuery: trimmed,
        }));
      }
    },
    [],
  );

  /**
   * 「下一个 / 上一个」：如果用户改过查询词（或还没有结果）就先重新查找，
   * 否则在现有结果里移动；未开启循环查找时到边界只提示、不绕回。
   */
  const next = useCallback(
    (delta: number) => {
      const current = stateRef.current;
      if (current.query.trim() !== current.resultQuery || current.hits.length === 0) {
        void run();
        return;
      }
      let index = current.index + delta;
      if (index < 0 || index >= current.hits.length) {
        if (!current.options.wrap) {
          setState((prev) => ({
            ...prev,
            status: delta > 0 ? "已到末尾（未开启循环查找）" : "已到开头（未开启循环查找）",
          }));
          return;
        }
        index = (index + current.hits.length) % current.hits.length;
      }
      jumpRef.current(current.hits[index]);
      setState((prev) => ({ ...prev, index, status: "" }));
    },
    [run],
  );

  const open = useCallback(() => setState((prev) => ({ ...prev, open: true })), []);
  const close = useCallback(
    () =>
      setState((prev) => ({
        ...prev,
        open: false,
        // 关掉面板只清查找结果；开关 / 范围 / 替换词留着（下次打开还是用户上次的选择）
        hits: [],
        index: -1,
        query: "",
        resultQuery: "",
        error: "",
        status: "",
      })),
    [],
  );
  const setQuery = useCallback(
    (query: string) => setState((prev) => ({ ...prev, query, status: "" })),
    [],
  );
  const setReplacement = useCallback(
    (replacement: string) => setState((prev) => ({ ...prev, replacement })),
    [],
  );
  const toggleReplace = useCallback(
    () => setState((prev) => ({ ...prev, replaceOpen: !prev.replaceOpen })),
    [],
  );
  const setScope = useCallback(
    (scope: SheetFindScope) => {
      if (stateRef.current.scope === scope) return;
      setState((prev) => ({ ...prev, scope, status: "" }));
      // 范围变了，结果必须重算（有查询词时立刻用**新范围**重查）
      if (stateRef.current.query.trim() !== "") void run(undefined, undefined, scope);
    },
    [run],
  );
  const toggleOption = useCallback(
    (key: keyof SheetFindOptions) => {
      if (key === "regex" && !supportsRegex) return; // xlsx：正则恒不可用
      const nextOptions = { ...stateRef.current.options, [key]: !stateRef.current.options[key] };
      setState((prev) => ({ ...prev, options: nextOptions, status: "" }));
      // 开关变了，结果必须重算（有查询词时立刻用**新开关**重查）
      if (stateRef.current.query.trim() !== "") void run(undefined, nextOptions);
    },
    [run, supportsRegex],
  );

  /**
   * 替换当前命中并跳到下一个（Excel / 编辑器惯例）。
   * 替换成功后把这一条从本地命中表里摘掉（它已经不匹配了），下标原地不动即指向下一个。
   */
  const replaceCurrent = useCallback(async () => {
    const current = stateRef.current;
    const replace = replaceCurrentRef.current;
    if (!replace) {
      setState((prev) => ({ ...prev, status: "当前文档不支持替换" }));
      return;
    }
    if (current.error) return; // 正则非法等错误：按钮本来就是禁用的
    if (current.query.trim() === "") {
      setState((prev) => ({ ...prev, status: "先输入查找内容" }));
      return;
    }
    const hit = current.index >= 0 ? current.hits[current.index] : undefined;
    if (!hit) {
      setState((prev) => ({ ...prev, status: "没有可替换的命中" }));
      return;
    }
    try {
      const count = await replace(current.replacement, {
        options: current.options,
        scope: current.scope,
        hits: current.hits,
        hit,
      });
      if (count > 0) {
        const hits = current.hits.filter((_, index) => index !== current.index);
        const index = hits.length === 0 ? -1 : Math.min(current.index, hits.length - 1);
        if (index >= 0) jumpRef.current(hits[index]);
        setState((prev) => ({ ...prev, hits, index, status: `已替换 ${count} 处` }));
      } else {
        setState((prev) => ({ ...prev, status: "这一处内容没有变化，未替换" }));
      }
    } catch (error) {
      setState((prev) => ({ ...prev, status: "", error: messageOf(error) }));
    }
  }, []);

  /** 全部替换：交给父组件一次做完（父组件负责记成**一个**撤销步），然后重查刷新命中 */
  const replaceAll = useCallback(async () => {
    const current = stateRef.current;
    const replace = replaceAllRef.current;
    if (!replace) {
      setState((prev) => ({ ...prev, status: "当前文档不支持替换" }));
      return;
    }
    if (current.error) return;
    if (current.query.trim() === "") {
      setState((prev) => ({ ...prev, status: "先输入查找内容" }));
      return;
    }
    if (current.hits.length === 0) {
      setState((prev) => ({ ...prev, status: "没有可替换的命中" }));
      return;
    }
    try {
      const count = await replace(current.replacement, {
        options: current.options,
        scope: current.scope,
        hits: current.hits,
        hit: null,
      });
      // 先重查（被替换掉的格不再命中），再落提示 —— 顺序反了会被 run 清掉
      await run();
      setState((prev) => ({ ...prev, status: `已替换 ${count} 处` }));
    } catch (error) {
      setState((prev) => ({ ...prev, status: "", error: messageOf(error) }));
    }
  }, [run]);

  /* Ctrl+F / Ctrl+H / F3 / Shift+F3 / Esc：窗口捕获、与焦点无关（保存弹窗关掉后仍能用） */
  useEffect(() => {
    if (!isActiveDoc) return;
    const onKeyDown = (event: KeyboardEvent) => {
      const mod = event.ctrlKey || event.metaKey;
      const target = event.target as HTMLElement | null;
      const inField = Boolean(target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA"));
      // 面板自己的输入框不算「别处的输入框」：在查找框里按 Ctrl+F / F3 仍然要生效
      const inPanel = Boolean(target && target.closest?.("[data-find-panel]"));
      if (mod && !event.altKey && event.key.toLowerCase() === "f") {
        if (inField && !inPanel) return;
        event.preventDefault();
        event.stopPropagation();
        open();
        return;
      }
      if (mod && !event.altKey && event.key.toLowerCase() === "h") {
        if (inField && !inPanel) return;
        event.preventDefault();
        event.stopPropagation();
        setState((prev) => ({ ...prev, open: true, replaceOpen: true }));
        return;
      }
      if (event.key === "F3" && (!inField || inPanel)) {
        event.preventDefault();
        event.stopPropagation();
        next(event.shiftKey ? -1 : 1);
        return;
      }
      if (event.key === "Escape") {
        setState((prev) => (prev.open ? { ...prev, open: false } : prev));
      }
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [isActiveDoc, next, open]);

  return useMemo(
    () => ({
      state,
      run,
      next,
      open,
      close,
      setQuery,
      toggleOption,
      setScope,
      setReplacement,
      toggleReplace,
      replaceCurrent,
      replaceAll,
    }),
    [
      state,
      run,
      next,
      open,
      close,
      setQuery,
      toggleOption,
      setScope,
      setReplacement,
      toggleReplace,
      replaceCurrent,
      replaceAll,
    ],
  );
}

export interface SheetFindBarProps {
  /** `useSheetFind` 的返回值 */
  api: SheetFindApi;
  /** 当前文档是否支持正则：xlsx 传 false（`.*` 置灰 + tooltip 说明） */
  supportsRegex?: boolean;
}

/**
 * 查找 / 替换栏：渲染**浮动面板**（fixed 定位，不占布局高度，也不会被网格的 overflow 裁掉）。
 *
 * 注意：面板用 `position: fixed`。若某个祖先带 `transform` / `filter` / `contain: paint`，
 * fixed 会退化成相对该祖先定位并被它裁剪 —— 请把面板挂在没有这些属性的层级上。
 */
export function SheetFindBar({ api, supportsRegex = true }: SheetFindBarProps) {
  return <SheetFindPanel api={api} supportsRegex={supportsRegex} />;
}

/** 浮动查找 / 替换面板 */
export function SheetFindPanel({ api, supportsRegex = true }: { api: SheetFindApi; supportsRegex?: boolean }) {
  const { state } = api;
  const queryRef = useRef<HTMLInputElement>(null);
  const replaceRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!state.open) return;
    queryRef.current?.focus();
    queryRef.current?.select();
  }, [state.open]);
  useEffect(() => {
    if (state.open && state.replaceOpen) replaceRef.current?.focus();
  }, [state.open, state.replaceOpen]);

  if (!state.open) return null;

  const total = state.hits.length;
  const current = state.index >= 0 ? state.index + 1 : 0;
  const stale = state.query.trim() !== state.resultQuery;
  const counter = state.busy
    ? "查找中…"
    : state.error
      ? "正则错误"
      : total > 0
        ? `${current}/${total}${state.capped ? "+" : ""}${stale ? " *" : ""}`
        : state.query
          ? "无结果"
          : "";
  // 有错误（正则非法 / 后端拒绝）或没有命中时禁用替换
  const canReplace = state.canReplace && !state.error && total > 0;
  const toggleClass = (active: boolean, disabled = false) =>
    `flex h-6 min-w-6 items-center justify-center rounded px-1 text-[11px] font-medium transition-colors ${
      disabled
        ? "cursor-not-allowed text-faint/50"
        : active
          ? "bg-accent-soft-strong font-semibold text-accent"
          : "text-muted hover:bg-hover hover:text-fg"
    }`;

  return (
    <div
      role="dialog"
      aria-label="查找和替换"
      data-find-panel="true"
      className="fixed right-6 top-16 z-50 w-[430px] rounded-[var(--radius)] border border-line bg-elevated p-2 shadow-[var(--shadow)]"
    >
      {/* 查找行 */}
      <div className="flex items-center gap-1">
        <button
          type="button"
          data-find-action="toggle-replace"
          title={state.replaceOpen ? "隐藏替换 (Ctrl+H)" : "显示替换 (Ctrl+H)"}
          onClick={api.toggleReplace}
          className={toggleClass(state.replaceOpen)}
        >
          <Icon name="chevron-down" size={13} className={state.replaceOpen ? "" : "-rotate-90"} />
        </button>
        <Icon name="search" size={14} className="shrink-0 text-faint" />
        <input
          ref={queryRef}
          data-find-input="query"
          value={state.query}
          onChange={(event) => api.setQuery(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              // 改过查询词（或还没有结果）就重新查找，否则在结果之间移动
              if (stale || total === 0) void api.run();
              else api.next(event.shiftKey ? -1 : 1);
            } else if (event.key === "Escape") {
              event.preventDefault();
              api.close();
            }
          }}
          placeholder="查找内容…"
          spellCheck={false}
          className="h-7 min-w-0 flex-1 rounded border border-line bg-input px-2 text-[12px] text-fg outline-none placeholder:text-faint focus:border-accent"
        />
        <span data-find-counter className="min-w-[52px] shrink-0 text-center text-[11px] text-faint">
          {counter}
        </span>
        <button
          type="button"
          data-find-action="prev"
          title="查找上一个 (Shift+F3)"
          onClick={() => api.next(-1)}
          className={toggleClass(false)}
        >
          <Icon name="arrow-up" size={14} />
        </button>
        <button
          type="button"
          data-find-action="next"
          title="查找下一个 (F3)"
          onClick={() => api.next(1)}
          className={toggleClass(false)}
        >
          <Icon name="arrow-down" size={14} />
        </button>
        <button
          type="button"
          data-find-action="close"
          title="关闭 (Esc)"
          onClick={api.close}
          className={toggleClass(false)}
        >
          <Icon name="x" size={14} />
        </button>
      </div>

      {/* 选项行 */}
      <div className="mt-1.5 flex items-center gap-1 pl-6">
        <button
          type="button"
          data-find-option="matchCase"
          title="区分大小写"
          onClick={() => api.toggleOption("matchCase")}
          className={toggleClass(state.options.matchCase)}
        >
          Aa
        </button>
        <button
          type="button"
          data-find-option="wholeCell"
          title="全字匹配（整格内容完全相同）"
          onClick={() => api.toggleOption("wholeCell")}
          className={toggleClass(state.options.wholeCell)}
        >
          <span className="underline decoration-dotted">W</span>
        </button>
        <button
          type="button"
          data-find-option="regex"
          title={supportsRegex ? "正则表达式" : "Excel 表格暂不支持正则查找（只有 CSV 支持）"}
          disabled={!supportsRegex}
          onClick={() => api.toggleOption("regex")}
          className={toggleClass(state.options.regex, !supportsRegex)}
        >
          .*
        </button>
        <button
          type="button"
          data-find-option="wrap"
          title="循环查找（到达末尾后回到开头）"
          onClick={() => api.toggleOption("wrap")}
          className={toggleClass(state.options.wrap)}
        >
          <Icon name="refresh" size={13} />
        </button>
        <div className="flex-1" />
        <span className="shrink-0 text-[11px] text-faint">范围</span>
        <button
          type="button"
          data-find-scope="sheet"
          title="在整张工作表里查找"
          onClick={() => api.setScope("sheet")}
          className={toggleClass(state.scope === "sheet")}
        >
          整表
        </button>
        <button
          type="button"
          data-find-scope="selection"
          title="只在当前选区里查找与替换"
          onClick={() => api.setScope("selection")}
          className={toggleClass(state.scope === "selection")}
        >
          选区
        </button>
      </div>

      {state.error ? (
        <div
          data-find-error
          className="mt-1.5 flex items-start gap-1 rounded bg-danger-soft px-2 py-1 text-[11px] text-danger"
        >
          <Icon name="alert-triangle" size={12} className="mt-0.5 shrink-0" />
          <span className="break-all">{state.error}</span>
        </div>
      ) : null}
      {!state.error && state.status ? (
        <div data-find-status className="mt-1.5 rounded bg-accent-soft px-2 py-1 text-[11px] text-accent">
          {state.status}
        </div>
      ) : null}

      {/* 替换行 */}
      {state.replaceOpen ? (
        <>
          <div className="mt-1.5 flex items-center gap-1">
            <span className="w-6 shrink-0" />
            <Icon name="replace" size={14} className="shrink-0 text-faint" />
            <input
              ref={replaceRef}
              data-find-input="replacement"
              value={state.replacement}
              onChange={(event) => api.setReplacement(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  void api.replaceCurrent();
                } else if (event.key === "Escape") {
                  event.preventDefault();
                  api.close();
                }
              }}
              placeholder="替换为…"
              spellCheck={false}
              className="h-7 min-w-0 flex-1 rounded border border-line bg-input px-2 text-[12px] text-fg outline-none placeholder:text-faint focus:border-accent"
            />
            <button
              type="button"
              data-find-action="replace"
              title="替换当前命中并跳到下一个 (Enter)"
              disabled={!canReplace}
              onClick={() => void api.replaceCurrent()}
              className={`h-7 shrink-0 rounded border border-line bg-elevated px-2.5 text-[11px] ${
                canReplace ? "text-fg hover:bg-hover" : "cursor-not-allowed text-faint/60"
              }`}
            >
              替换
            </button>
            <button
              type="button"
              data-find-action="replace-all"
              title="替换全部命中（父组件记成一个撤销步）"
              disabled={!canReplace}
              onClick={() => void api.replaceAll()}
              className={`h-7 shrink-0 rounded border border-transparent px-2.5 text-[11px] font-semibold ${
                canReplace ? "bg-accent text-accent-fg hover:opacity-90" : "cursor-not-allowed bg-hover text-faint/60"
              }`}
            >
              全部替换
            </button>
          </div>
          <div className="mt-1 pl-6 text-[11px] text-faint">
            替换只作用于<b>当前工作表</b>
            {state.scope === "selection" ? "的当前选区" : ""}；跨工作表的命中只能查找、不能替换。
          </div>
        </>
      ) : null}
    </div>
  );
}
