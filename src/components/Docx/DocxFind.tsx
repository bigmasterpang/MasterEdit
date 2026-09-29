/**
 * 文档查找面板（`Ctrl+F`）。
 *
 * 交互风格对齐表格的 `SheetFindBar`（Notepad++ 风格的浮动面板），但**语义更简单**：
 * 文档是线性的，命中就是「第几个块」，所以只需要「上一个 / 下一个 + 命中计数」，
 * 没有替换、没有范围与正则（后端 `document_find` 只有 `matchCase`）。
 *
 * 命中来源是后端 `document_find`（全文档扫描，块级命中，最多 500 条），
 * 跳转由父组件负责（滚动到块 + 加高亮）。
 */
import { useCallback, useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import type { DocFindHit } from "../../types";
import { Icon } from "../common/Icon";

/** 输入停顿多久才发查询（文档查找要全篇扫描，逐字符发请求太浪费） */
const SEARCH_DEBOUNCE_MS = 150;
/** 命中数上限（与后端 `MAX_FIND_HITS` 对齐） */
export const FIND_HIT_LIMIT = 500;

export interface DocxFindApi {
  open: boolean;
  query: string;
  matchCase: boolean;
  hits: DocFindHit[];
  /** 当前命中下标（-1 = 没有命中） */
  index: number;
  busy: boolean;
  error: string;
  /** 一次性提示（没有命中 / 已到末尾…） */
  status: string;
  openPanel(): void;
  close(): void;
  setQuery(value: string): void;
  setMatchCase(value: boolean): void;
  next(): void;
  prev(): void;
}

export interface UseDocxFindOptions {
  /** 文档路径（null = 还没有文档） */
  path: string | null;
  /** 是否为当前激活文档（双栏时只有激活的那个响应 Ctrl+F） */
  active: boolean;
  /** 跳到某条命中（父组件滚动 + 高亮） */
  jump(hit: DocFindHit): void;
}

export function useDocxFind({ path, active, jump }: UseDocxFindOptions): DocxFindApi {
  const [open, setOpen] = useState(false);
  const [query, setQueryState] = useState("");
  const [matchCase, setMatchCaseState] = useState(false);
  const [hits, setHits] = useState<DocFindHit[]>([]);
  const [index, setIndex] = useState(-1);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  /** 当前结果对应的查询词：与 query 不一致说明用户改了内容、还没重新查找 */
  const resultQueryRef = useRef("");
  const inputRef = useRef<HTMLInputElement | null>(null);

  const clearResults = useCallback((message = "") => {
    setHits([]);
    setIndex(-1);
    setStatus(message);
  }, []);

  /** 执行查找（返回命中列表，供"回车立即搜"复用） */
  const runFind = useCallback(
    async (value: string, caseSensitive: boolean): Promise<DocFindHit[]> => {
      const keyword = value.trim();
      if (!path || !keyword) {
        resultQueryRef.current = "";
        clearResults("");
        return [];
      }
      setBusy(true);
      setError("");
      try {
        const result = await invoke<DocFindHit[]>("document_find", {
          path,
          query: keyword,
          matchCase: caseSensitive,
        });
        resultQueryRef.current = keyword;
        setHits(result);
        setIndex(result.length > 0 ? 0 : -1);
        setStatus(
          result.length === 0
            ? "没有找到匹配的内容"
            : result.length >= FIND_HIT_LIMIT
              ? `命中较多，只显示前 ${FIND_HIT_LIMIT} 条`
              : "",
        );
        return result;
      } catch (reason) {
        setError(typeof reason === "string" && reason ? reason : String(reason));
        clearResults("");
        return [];
      } finally {
        setBusy(false);
      }
    },
    [path, clearResults],
  );

  /* 输入防抖查找 */
  useEffect(() => {
    if (!open) return;
    if (!query.trim()) {
      resultQueryRef.current = "";
      clearResults("");
      return;
    }
    const timer = window.setTimeout(() => {
      void runFind(query, matchCase);
    }, SEARCH_DEBOUNCE_MS);
    return () => window.clearTimeout(timer);
  }, [query, matchCase, open, runFind, clearResults]);

  /* 命中变化 → 通知父组件滚动并高亮（在查找效果里同步做，跳转永远跟着当前命中） */
  const jumpRef = useRef(jump);
  jumpRef.current = jump;
  useEffect(() => {
    if (!open) return;
    const hit = index >= 0 ? hits[index] : undefined;
    if (hit) jumpRef.current(hit);
  }, [hits, index, open]);

  const step = useCallback(
    (delta: number) => {
      if (hits.length === 0) return;
      const next = (index + delta + hits.length) % hits.length;
      setIndex(next);
      setStatus(
        hits.length > 1 && ((delta > 0 && next === 0) || (delta < 0 && next === hits.length - 1))
          ? delta > 0
            ? "已回到第一条"
            : "已到最后一条"
          : "",
      );
    },
    [hits.length, index],
  );

  const openPanel = useCallback(() => {
    setOpen(true);
    // 打开后聚焦输入框（并选中已有内容，方便直接改写查询词）
    window.setTimeout(() => {
      inputRef.current?.focus();
      inputRef.current?.select();
    }, 0);
  }, []);

  const close = useCallback(() => setOpen(false), []);

  const setQuery = useCallback((value: string) => setQueryState(value), []);
  const setMatchCase = useCallback((value: boolean) => setMatchCaseState(value), []);

  /* 快捷键：Ctrl+F 打开、Esc 关闭、F3 / Shift+F3 上下一条 */
  useEffect(() => {
    if (!active) return;
    const onKeyDown = (event: KeyboardEvent) => {
      const mod = event.ctrlKey || event.metaKey;
      if (mod && !event.altKey && event.key.toLowerCase() === "f") {
        event.preventDefault();
        event.stopPropagation();
        openPanel();
        return;
      }
      if (event.key === "F3") {
        event.preventDefault();
        if (!open) openPanel();
        else step(event.shiftKey ? -1 : 1);
        return;
      }
      if (event.key === "Escape" && open) {
        event.preventDefault();
        close();
      }
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [active, open, openPanel, close, step]);

  /* 文档换了：结果作废 */
  useEffect(() => {
    resultQueryRef.current = "";
    clearResults("");
    setQueryState("");
  }, [path, clearResults]);

  return {
    open,
    query,
    matchCase,
    hits,
    index,
    busy,
    error,
    status,
    openPanel,
    close,
    setQuery,
    setMatchCase,
    next: () => step(1),
    prev: () => step(-1),
  };
}

/** 面板 UI：绝对定位在内容区右上角 */
export function DocxFindBar({ api }: { api: DocxFindApi }) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    if (api.open) inputRef.current?.focus();
  }, [api.open]);

  if (!api.open) return null;

  const total = api.hits.length;
  const position = api.index >= 0 ? api.index + 1 : 0;

  return (
    <div
      data-docx-find="true"
      className="absolute right-3 top-3 z-30 w-[300px] rounded-lg border border-line bg-elevated/95 p-2 shadow-lg backdrop-blur"
    >
      <div className="flex items-center gap-1">
        <Icon name="search" size={13} className="shrink-0 text-muted" />
        <input
          ref={inputRef}
          data-docx-find-input="true"
          value={api.query}
          onChange={(event) => api.setQuery(event.target.value)}
          onKeyDown={(event) => {
            if (event.key !== "Enter") return;
            event.preventDefault();
            if (event.shiftKey) api.prev();
            else api.next();
          }}
          placeholder="在文档中查找"
          spellCheck={false}
          className="h-6 min-w-0 flex-1 rounded border border-line bg-app px-1.5 text-[12px] text-fg outline-none focus:border-accent"
        />
        <button
          type="button"
          data-docx-find-case="true"
          aria-pressed={api.matchCase}
          title="区分大小写"
          onClick={() => api.setMatchCase(!api.matchCase)}
          className={`h-6 rounded border px-1.5 text-[11px] ${
            api.matchCase ? "border-accent bg-accent/15 text-accent" : "border-line text-muted hover:bg-hover"
          }`}
        >
          Aa
        </button>
        <button
          type="button"
          data-docx-find-prev="true"
          title="上一个（Shift+F3）"
          disabled={total === 0}
          onClick={() => api.prev()}
          className="flex h-6 w-6 items-center justify-center rounded border border-line text-muted hover:bg-hover disabled:opacity-40"
        >
          <Icon name="chevron-up" size={13} />
        </button>
        <button
          type="button"
          data-docx-find-next="true"
          title="下一个（F3 / Enter）"
          disabled={total === 0}
          onClick={() => api.next()}
          className="flex h-6 w-6 items-center justify-center rounded border border-line text-muted hover:bg-hover disabled:opacity-40"
        >
          <Icon name="chevron-down" size={13} />
        </button>
        <button
          type="button"
          data-docx-find-close="true"
          title="关闭（Esc）"
          onClick={() => api.close()}
          className="flex h-6 w-6 items-center justify-center rounded border border-line text-muted hover:bg-hover"
        >
          <Icon name="x" size={13} />
        </button>
      </div>
      <div className="mt-1 flex items-center justify-between text-[11px] text-muted">
        <span data-docx-find-count="true">
          {total === 0 ? (api.busy ? "查找中…" : "0 条命中") : `${position}/${total}`}
        </span>
        <span className="truncate text-faint">{api.error || api.status}</span>
      </div>
    </div>
  );
}
