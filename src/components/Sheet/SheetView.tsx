import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { SheetGrid } from "./SheetGrid";
import { createDelimitedTable, detectDelimiter } from "../../utils/delimited";
import { useAppStore } from "../../stores/appStore";
import { extName, isSpreadsheetDoc } from "../../utils/filePath";
import type { SheetCell, SpreadsheetInfo, SpreadsheetWindow } from "../../types";

/** 单次向 Rust 请求的行数（后端还有单元格数上限，会自动收窄） */
const WINDOW_ROWS = 300;

/** 列序号 → Excel 列标（0 → A，26 → AA） */
function columnLabel(index: number): string {
  let n = index;
  let label = "";
  while (n >= 0) {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  }
  return label;
}

/** 单元格引用（0 起行列 → B3） */
function cellRef(row: number, col: number): string {
  return `${columnLabel(col)}${row + 1}`;
}

/** 内容防抖：分屏编辑 CSV 时每次按键都会触发解析，需要节流 */
function useDebounced<T>(value: T, delay = 200): T {
  const [debounced, setDebounced] = useState(value);
  useEffect(() => {
    const timer = window.setTimeout(() => setDebounced(value), delay);
    return () => window.clearTimeout(timer);
  }, [value, delay]);
  return debounced;
}

interface SheetViewProps {
  docId: string;
}

/**
 * 表格视图入口：按文档类型分派到「分隔符文本表格（CSV/TSV）」或「工作簿只读表格（xlsx/xls/ods）」。
 * 由 DocView 懒加载，未打开表格文档时不会进入首屏包。
 */
export function SheetView({ docId }: SheetViewProps) {
  const isWorkbook = useAppStore((s) => isSpreadsheetDoc(s.docs.find((d) => d.id === docId) ?? null));
  return isWorkbook ? <WorkbookSheet docId={docId} /> : <DelimitedSheet docId={docId} />;
}

/* ------------------------------------------------------------------ */
/* 公共外壳：单元格内容栏 + 表格 + 可选工作表标签                        */
/* ------------------------------------------------------------------ */

interface SheetFrameProps {
  /** 左上角显示的定位信息（如 B3） */
  reference: string;
  /** 单元格内容（公式优先） */
  content: string;
  readOnly?: boolean;
  /** 企业透明加密文档（已在内存中解密） */
  encrypted?: boolean;
  /** 右侧附加信息（行数、截断提示等） */
  hint?: string;
  tabs?: React.ReactNode;
  children: React.ReactNode;
}

function SheetFrame({ reference, content, readOnly, encrypted, hint, tabs, children }: SheetFrameProps) {
  return (
    <div className="flex h-full min-w-0 flex-col bg-app">
      {/* 单元格内容栏（只读） */}
      <div className="flex h-7 shrink-0 items-center gap-2 border-b border-line bg-panel px-2 text-[12px]">
        <span className="w-14 shrink-0 truncate font-medium text-faint" title="当前单元格">
          {reference || "—"}
        </span>
        <span className="min-w-0 flex-1 truncate text-fg" title={content}>
          {content}
        </span>
        {hint ? <span className="shrink-0 text-faint">{hint}</span> : null}
        {encrypted ? (
          <span
            className="shrink-0 rounded bg-accent/10 px-1.5 py-0.5 text-[11px] text-accent"
            title="企业加密文档：已在内存中解密查看，不会写出明文副本"
          >
            已解密
          </span>
        ) : null}
        {readOnly ? (
          <span className="shrink-0 rounded bg-warning/10 px-1.5 py-0.5 text-[11px] text-warning" title="表格仅支持查看，不写回原文件">
            只读
          </span>
        ) : null}
      </div>

      <div className="relative min-h-0 flex-1">{children}</div>

      {tabs ? (
        <div className="flex h-7 shrink-0 items-center gap-1 overflow-x-auto border-t border-line bg-panel px-2">
          {tabs}
        </div>
      ) : null}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* CSV / TSV：复用文本管线内容，前端解析                                */
/* ------------------------------------------------------------------ */

function DelimitedSheet({ docId }: { docId: string }) {
  const content = useAppStore((s) => s.docs.find((d) => d.id === docId)?.content ?? "");
  const filePath = useAppStore((s) => s.docs.find((d) => d.id === docId)?.filePath ?? null);
  const [active, setActive] = useState<{ row: number; col: number } | null>(null);
  const debounced = useDebounced(content, 200);

  /**
   * 惰性行索引：只单遍扫描记录行首偏移（十几 MB / 十几万行只需几十毫秒，
   * 额外内存只有一张偏移表），行内容在网格渲染到该行时才解析。
   * 因此没有行数上限，也不会像一次性解析那样为每个单元格建对象。
   */
  const table = useMemo(() => {
    const ext = extName(filePath ?? "");
    return createDelimitedTable(debounced, { delimiter: detectDelimiter(debounced, ext), ext });
  }, [debounced, filePath]);

  // getCell 的引用随索引变化，保证网格能拿到最新数据
  const getCell = useCallback(
    (row: number, col: number): SheetCell | undefined => table.rowAt(row)?.[col],
    [table],
  );

  const activeCell = active && active.row < table.rows ? active : null;
  const activeValue = activeCell ? (table.rowAt(activeCell.row)?.[activeCell.col]?.v ?? "") : "";
  const delimiterLabel = table.delimiter === "\t" ? "制表符" : table.delimiter;
  const hint = `${table.rows} 行${table.truncatedCols ? " · 列已截断" : ""}`;

  if (table.rows === 0) {
    return (
      <SheetFrame reference="" content="" hint="空表格">
        <div className="flex h-full items-center justify-center text-[12px] text-muted">
          空表格（没有可显示的行）
        </div>
      </SheetFrame>
    );
  }

  return (
    <SheetFrame
      reference={activeCell ? cellRef(activeCell.row, activeCell.col) : ""}
      content={activeValue}
      hint={hint}
      tabs={
        <span className="text-[11px] text-faint" title="自动识别的分隔符">
          分隔符：{delimiterLabel} · 列 {table.cols}
        </span>
      }
    >
      <SheetGrid
        rows={table.rows}
        cols={table.cols}
        getCell={getCell}
        /* CSV 首行通常是表头：冻结它，长表滚动时仍能看到列名 */
        headerRow
        activeCell={activeCell}
        onActiveCell={(row, col) => setActive({ row, col })}
      />
    </SheetFrame>
  );
}

/* ------------------------------------------------------------------ */
/* xlsx / xls / ods：Rust 侧解析，按行窗口按需加载（只读）               */
/* ------------------------------------------------------------------ */

function WorkbookSheet({ docId }: { docId: string }) {
  const filePath = useAppStore((s) => s.docs.find((d) => d.id === docId)?.filePath ?? null);
  // 外部改动由文件监听更新 modifiedAt：把它作为依赖即可自动重新解析
  const modifiedAt = useAppStore((s) => s.docs.find((d) => d.id === docId)?.modifiedAt ?? 0);
  const [info, setInfo] = useState<SpreadsheetInfo | null>(null);
  const [sheetIndex, setSheetIndex] = useState(0);
  const [dim, setDim] = useState<{ rows: number; cols: number } | null>(null);
  const [active, setActive] = useState<{ row: number; col: number } | null>(null);
  const [truncated, setTruncated] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  /** 已加载的行：行号 → 单元格（稀疏缓存，滚动到哪加载到哪） */
  const rowsRef = useRef<Map<number, SheetCell[]>>(new Map());
  /** 正在请求的区间，避免重复请求 */
  const pendingRef = useRef<Array<[number, number]>>([]);
  const [version, setVersion] = useState(0);

  const sheetName = info?.sheets[sheetIndex]?.name ?? null;

  /* 读取工作表列表 */
  useEffect(() => {
    if (!filePath) return;
    let cancelled = false;
    setLoading(true);
    setError(null);
    void (async () => {
      try {
        const result = await invoke<SpreadsheetInfo>("spreadsheet_info", { path: filePath });
        if (cancelled) return;
        setInfo(result);
        setSheetIndex(0);
        if (result.sheets.length === 0) {
          setError("该工作簿中没有工作表");
        }
      } catch (err) {
        if (!cancelled) setError(String(err));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [filePath, modifiedAt]);

  /* 切换工作表：清空行缓存与选中状态 */
  useEffect(() => {
    rowsRef.current.clear();
    pendingRef.current = [];
    setDim(null);
    setActive(null);
    setTruncated(false);
    setVersion(0);
  }, [sheetIndex, filePath]);

  /** 请求 [start, end) 区间内尚未缓存的行 */
  const loadRange = useCallback(
    (start: number, end: number) => {
      if (!filePath || !sheetName) return;
      const isPending = (from: number, to: number) =>
        pendingRef.current.some(([s, e]) => s <= from && e >= to);

      let cursor = Math.max(0, start);
      while (cursor < end) {
        if (rowsRef.current.has(cursor)) {
          cursor += 1;
          continue;
        }
        const requestEnd = Math.min(end, cursor + WINDOW_ROWS);
        if (!isPending(cursor, requestEnd)) {
          const from = cursor;
          const to = requestEnd;
          pendingRef.current.push([from, to]);
          void invoke<SpreadsheetWindow>("spreadsheet_rows", {
            path: filePath,
            sheet: sheetName,
            start: from,
            count: to - from,
          })
            .then((win) => {
              win.cells.forEach((cells, index) => rowsRef.current.set(win.start + index, cells));
              setDim({ rows: win.rows, cols: win.cols });
              if (win.truncated) setTruncated(true);
              setVersion((v) => v + 1);
            })
            .catch((err) => setError(String(err)))
            .finally(() => {
              pendingRef.current = pendingRef.current.filter(([s, e]) => !(s === from && e === to));
            });
        }
        cursor = requestEnd;
      }
    },
    [filePath, sheetName],
  );

  // 依赖 version 让 getCell 身份随数据变化，网格（可能被 memo）才会重绘
  const getCell = useCallback(
    (row: number, col: number): SheetCell | undefined => rowsRef.current.get(row)?.[col],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [version],
  );

  /* 首次进入工作表时主动拉第一批数据：
     此时网格还不知道总行数（rows=0），它上报的可视区间为空，无法驱动加载 */
  useEffect(() => {
    if (!sheetName || rowsRef.current.size > 0) return;
    loadRange(0, WINDOW_ROWS);
  }, [sheetName, loadRange]);

  const onViewport = useCallback(
    // SheetGrid 的 endRow 为开区间，直接透传即可
    (startRow: number, endRow: number) => loadRange(startRow, endRow),
    [loadRange],
  );

  const activeValue = active ? (rowsRef.current.get(active.row)?.[active.col]?.v ?? "") : "";
  const activeFormula = active ? rowsRef.current.get(active.row)?.[active.col]?.f : undefined;

  if (!filePath) {
    return <SheetFrame reference="" content=""><div /></SheetFrame>;
  }
  if (error) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-2 p-6 text-center">
        <div className="text-[13px] font-medium text-fg/80">无法打开表格</div>
        <div className="max-w-[420px] text-[12px] leading-relaxed text-muted">{error}</div>
        <div className="text-[11px] text-faint">
          提示：受密码保护或结构异常的工作簿暂不支持；可右键标签选择「用系统默认程序打开」。
        </div>
      </div>
    );
  }
  if (loading || !info) {
    return (
      <div className="flex h-full items-center justify-center text-[12px] text-muted">正在解析工作簿…</div>
    );
  }

  return (
    <SheetFrame
      reference={active ? cellRef(active.row, active.col) : ""}
      content={activeFormula ?? activeValue}
      readOnly
      encrypted={info.encrypted}
      hint={dim ? `${dim.rows} 行 × ${dim.cols} 列${truncated ? " · 已截断" : ""}` : "正在加载…"}
      tabs={
        <>
          {info.sheets.map((sheet, index) => (
            <button
              key={`${sheet.name}-${index}`}
              type="button"
              onClick={() => setSheetIndex(index)}
              title={sheet.name}
              className={`max-w-[180px] shrink-0 truncate rounded px-2 py-0.5 text-[11px] transition-colors ${
                index === sheetIndex
                  ? "bg-accent/15 font-medium text-accent"
                  : "text-muted hover:bg-hover hover:text-fg"
              }`}
            >
              {sheet.name}
            </button>
          ))}
        </>
      }
    >
      <SheetGrid
        rows={dim?.rows ?? 0}
        cols={dim?.cols ?? 0}
        getCell={getCell}
        onViewport={onViewport}
        activeCell={active}
        onActiveCell={(row, col) => setActive({ row, col })}
      />
    </SheetFrame>
  );
}
