import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { SheetGrid, type SheetRange } from "./SheetGrid";
import { createDelimitedTable, detectDelimiter } from "../../utils/delimited";
import { applyCellEditsToDelimited } from "../../utils/sheetEdits";
import {
  commitHistoryStep,
  mapCellBatch,
  mapCellKey,
  popHistoryStep,
  popLayoutOp,
  pushCellStep,
  pushLayoutOpAgain,
  pushStructureStep,
  returnHistoryStep,
  sheetHistoryDepth,
  type CellStep,
  type LayoutOp,
} from "../../utils/sheetHistory";
import { useSheetUndoRedo } from "../../hooks/useSheetUndoRedo";
import { useSheetLayout } from "../../hooks/useSheetLayout";
import { onSheetCommand } from "../../utils/sheetCommands";
import {
  SheetFindBar,
  useSheetFind,
  type SheetFindContext,
  type SheetFindHit,
  type SheetFindOptions,
  type SheetFindScope,
} from "./SheetFind";
import { buildCellMatcher, compileQuery, inSelectionRange, replaceCellText } from "../../utils/sheetReplace";
import { FormulaSuggestPanel, useFormulaSuggest } from "./FormulaSuggest";
import {
  extractFormulaReferences,
  matchFormulaFunctions,
  selectionReference,
} from "../../utils/formulaFunctions";
import { useAppStore } from "../../stores/appStore";
import { askConfirm, askForm, showMessage } from "../../stores/dialogStore";
import { ContextMenu } from "../common/ContextMenu";
import { extName, isSpreadsheetDoc } from "../../utils/filePath";
import type {
  EvalResult,
  RangeStats,
  SheetCell,
  SheetCellType,
  SheetEdit,
  SheetEditKind,
  ShadowEditResult,
  StructureResult,
  SpreadsheetInfo,
  SpreadsheetWindow,
} from "../../types";

/** 单次向 Rust 请求的行数（后端还有单元格数上限，会自动收窄） */
const WINDOW_ROWS = 300;

/** 无待提交编辑时的稳定空数组（避免每次渲染产生新引用） */
const EMPTY_EDITS: SheetEdit[] = [];

/** 单元格编辑在 Map 里的键 */
function editKey(sheet: string, row: number, col: number): string {
  return `${sheet}|${row}|${col}`;
}

/**
 * 由输入文本推断写回方式（与 Excel 手感一致）。
 * 与 Excel 的唯一差别：前导零的数字（编号、电话等）按文本写入，避免 007 变成 7。
 */
function inferEditKind(text: string, original?: SheetCellType): SheetEditKind {
  const trimmed = text.trim();
  if (trimmed === "") return "empty";
  if (trimmed.startsWith("=")) return "formula";
  if (/^(true|false)$/i.test(trimmed)) return "bool";
  if (original === "date") return "date";
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?$/.test(trimmed)) return "date";
  if (/^[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
    return /^[+-]?0\d/.test(trimmed) ? "text" : "number";
  }
  return "text";
}

/**
 * 撤销栈里的一步：记录一格编辑前后的状态。
 *
 * before 必须是**文件当前状态的显式编辑**（哪怕那一格本来是空的，也要写成 `kind: "empty"`）：
 * 保存会把待提交编辑清空，撤销时靠 before 重新写回一条待提交编辑来恢复界面；
 * 若 before 为 null（早期版本对空单元格就是这样），保存后 Ctrl+Z 会退化成「删掉一条不存在的编辑」，
 * 表现为撤销计数在动、内容却毫无变化。
 */
interface EditStep {
  key: string;
  before: SheetEdit;
  after: SheetEdit;
}

/* ------------------------------------------------------------------ */
/* 选区：序列化 / 剪贴板解析 / 统计                                     */
/* ------------------------------------------------------------------ */

/** 选区里能选中的最大单元格数（清空 / 粘贴的批量上限，超过先确认） */
const BULK_LIMIT = 5000;

/**
 * 网格的最小可编辑区域。
 * 空工作表（新建的表）与数据很少的表如果只渲染已用区域，用户根本点不进去、无法录入，
 * 因此始终留出 1000×50 的可编辑空间（滚动是虚拟化的，不影响性能），同时也能在数据下方追加。
 */
const GRID_MIN_ROWS = 1000;
const GRID_MIN_COLS = 50;

/** 取单元格文本（两个路径的取值方式不同，用回调注入） */
type CellReader = (row: number, col: number) => SheetCell | undefined;

/** 把选区序列化成 TSV（Excel / 表格软件通用的剪贴板格式，行尾用 CRLF） */
function rangeToTsv(range: SheetRange, read: CellReader): string {
  const lines: string[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const cells: string[] = [];
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      cells.push(read(row, col)?.v ?? "");
    }
    lines.push(cells.join("\t"));
  }
  return lines.join("\r\n");
}

/** 解析剪贴板里的表格文本：按行拆分、按制表符分列（首尾空行忽略） */
function parseClipboardTable(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  while (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
  return lines.map((line) => line.split("\t"));
}

/** 文本转数字（纯数字文本也参与统计：CSV 导入后数字常常是文本） */
function textToNumber(text: string): number | null {
  const trimmed = text.trim().replace(/,/g, "");
  if (trimmed === "" || !/^[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) return null;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

interface RangeSummary {
  cells: number;
  nonEmpty: number;
  numeric: number;
  sum: number;
  average: number | null;
  min: number | null;
  max: number | null;
  /** 选区是否只覆盖了已加载的数据（未加载时统计不完整，界面要标注） */
  partial: boolean;
}

/** 统计选区：求和 / 平均 / 数值个数 / 最大最小（只统计数值，文本数字也算） */
function summarizeRange(range: SheetRange, read: CellReader): RangeSummary {
  let cells = 0;
  let nonEmpty = 0;
  let numeric = 0;
  let sum = 0;
  let min: number | null = null;
  let max: number | null = null;
  let partial = false;
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      cells += 1;
      const cell = read(row, col);
      if (cell === undefined) {
        partial = true;
        continue;
      }
      if (cell.t === "empty" && cell.v === "") continue;
      nonEmpty += 1;
      const value = cell.t === "number" ? textToNumber(cell.v) : textToNumber(cell.v);
      if (value === null) continue;
      numeric += 1;
      sum += value;
      if (min === null || value < min) min = value;
      if (max === null || value > max) max = value;
    }
  }
  return {
    cells,
    nonEmpty,
    numeric,
    sum,
    average: numeric > 0 ? sum / numeric : null,
    min,
    max,
    partial,
  };
}

/** 数字显示：最多 4 位小数，去掉多余的 0 */
function formatStat(value: number | null): string {
  if (value === null) return "—";
  const rounded = Math.round(value * 10000) / 10000;
  return String(rounded);
}

/** 选区信息条：显示选区尺寸、统计值与提示（两个表格路径共用） */
function SelectionSummary({ range, summary }: { range: SheetRange | null; summary: RangeSummary | null }) {
  if (!range || !summary) {
    return <span className="text-[11px] text-faint">拖动可框选一片区域，Ctrl+C 复制、Delete 清空</span>;
  }
  const cols = range.endCol - range.startCol + 1;
  const rows = range.endRow - range.startRow + 1;
  return (
    <span className="flex items-center gap-2 text-[11px] text-muted">
      <span className="text-faint">
        选中 {rows}×{cols}
      </span>
      {summary.numeric > 0 ? (
        <>
          <span title="求和">求和 {formatStat(summary.sum)}</span>
          <span title="平均值">平均 {formatStat(summary.average)}</span>
          <span title="数值个数 / 非空个数">
            数值 {summary.numeric}/{summary.nonEmpty}
          </span>
          <span title="最小值">最小 {formatStat(summary.min)}</span>
          <span title="最大值">最大 {formatStat(summary.max)}</span>
        </>
      ) : (
        <span className="text-faint">非空 {summary.nonEmpty} 格（无数字可统计）</span>
      )}
      {summary.partial ? (
        <span className="text-warning" title="选区超出已加载的数据窗口，统计仅覆盖已加载部分">
          *部分
        </span>
      ) : null}
    </span>
  );
}

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
  /** 内容栏右侧的提示文字（也可以是带样式的节点，例如公式计算错误） */
  hint?: React.ReactNode;
  /** 可编辑时的内容栏输入框（不传则只读显示文本） */
  bar?: {
    value: string;
    onChange: (value: string) => void;
    onCommit: () => void;
    onCancel: () => void;
    /** 是否正在编辑（决定左侧「取消 / 接受」按钮可用，与 Excel 一致） */
    editing?: boolean;
    /** 点左侧 fx「插入公式」：父组件把内容栏切到公式编辑状态（通常置为 `=`） */
    onStartFormula?: () => void;
  };
  /** 查找栏：内联渲染在内容栏同一行右侧，不额外占一行 */
  findBar?: React.ReactNode;
  /** 当前工作表已用行列数：用于公式补全里的整列/整行/已用区域引用候选 */
  usedRows?: number;
  usedCols?: number;
  /** 当前选区的引用文本（如 `B2:D10` 或整列 `B:B`）：编辑公式时可一键插入 */
  selectionRef?: string;
  /** 把「写入区域引用」的能力交给父组件（网格拖选后回填公式） */
  onBarApi?: (api: { insertReference: (reference: string) => void } | null) => void;
  /** 内容栏是否正在编辑公式：父组件据此决定网格是否开启"鼠标拾取范围" */
  onPickModeChange?: (active: boolean) => void;
  tabs?: React.ReactNode;
  /** 底部条右侧的状态（选区统计、加载提示等） */
  statusRight?: React.ReactNode;
  children: React.ReactNode;
}

function SheetFrame({ reference, content, readOnly, encrypted, hint, bar, findBar, usedRows, usedCols, selectionRef, onBarApi, onPickModeChange, tabs, statusRight, children }: SheetFrameProps) {
  const barInputRef = useRef<HTMLInputElement | null>(null);
  // 内容栏里输入 `=` 时给出常用函数提示（只做名称/用法补全，不计算公式）
  const suggest = useFormulaSuggest({
    enabled: Boolean(bar),
    value: bar?.value ?? "",
    onChange: (next) => bar?.onChange(next),
    inputRef: barInputRef,
    context: { usedRows, usedCols },
  });

  /**
   * 把「写入区域引用」的能力交给父组件（网格里拖选范围后回填到公式里）。
   * API 对象用 ref 保持稳定身份，内部始终调用最新的闭包，避免每次渲染都回调父组件。
   */
  const apiRef = useRef<{ insertReference: (reference: string) => void }>({
    insertReference: () => undefined,
  });
  apiRef.current.insertReference = (reference: string) => suggest.insertReferenceText(reference);
  useEffect(() => {
    onBarApi?.(apiRef.current);
    return () => onBarApi?.(null);
  }, [onBarApi]);

  /** 公式编辑状态变化时通知父组件（决定网格是否开启拾取模式） */
  useEffect(() => {
    onPickModeChange?.(suggest.editingFormula);
  }, [onPickModeChange, suggest.editingFormula]);
  return (
    <div className="flex h-full min-w-0 flex-col bg-app">
      {/* 单元格内容栏（只读） */}
      <div className="flex h-7 shrink-0 items-center gap-2 border-b border-line bg-panel px-2 text-[12px]">
        <span className="w-14 shrink-0 truncate font-medium text-faint" title="当前单元格">
          {reference || "—"}
        </span>
        {bar ? (
          <>
            {/* 内容栏左侧三个按钮，与 Excel 一致：取消 / 接受 / 插入公式 */}
            <button
              type="button"
              onClick={bar.onCancel}
              disabled={!bar.editing}
              title="取消编辑 (Esc)"
              className="shrink-0 rounded px-1 text-muted hover:bg-hover hover:text-fg disabled:opacity-30"
            >
              ✕
            </button>
            <button
              type="button"
              onClick={bar.onCommit}
              disabled={!bar.editing}
              title="接受并写入单元格 (Enter)"
              className="shrink-0 rounded px-1 text-muted hover:bg-hover hover:text-fg disabled:opacity-30"
            >
              ✓
            </button>
            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                if (!bar.editing) bar.onStartFormula?.();
                barInputRef.current?.focus();
              }}
              title="插入公式（以 = 开头，会提示常用函数）"
              className="shrink-0 rounded px-1 font-serif italic text-muted hover:bg-hover hover:text-fg"
            >
              fx
            </button>
            <input
              ref={barInputRef}
              value={bar.value}
              onChange={(event) => bar.onChange(event.target.value)}
              onSelect={() => suggest.refresh()}
              onBlur={() => suggest.reset()}
              onKeyDown={(event) => {
                // 提示面板打开时先让它消费 ↑↓ / Tab / Enter / Esc
                if (suggest.handleKeyDown(event)) return;
                if (event.key === "Enter") {
                  event.preventDefault();
                  bar.onCommit();
                } else if (event.key === "Escape") {
                  event.preventDefault();
                  bar.onCancel();
                }
              }}
              spellCheck={false}
              placeholder="输入内容后回车写入单元格（= 开头为公式，会提示常用函数）"
              title="单元格内容 / 公式：回车写入待提交编辑，Ctrl+S 保存到文件"
              className="min-w-0 flex-1 rounded border border-line bg-app px-1.5 py-0.5 font-mono text-[12px] text-fg outline-none focus:border-accent"
            />
          </>
        ) : (
          <span className="min-w-0 flex-1 truncate text-fg" title={content}>
            {content}
          </span>
        )}
        <FormulaSuggestPanel state={suggest.state} onPick={(fn) => suggest.accept(fn)} />
        {bar && selectionRef ? (
          <button
            type="button"
            title={`在光标处插入选区引用 ${selectionRef}`}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => suggest.insertAtCaret(selectionRef)}
            className="shrink-0 rounded border border-line px-1.5 py-0.5 font-mono text-[11px] text-muted hover:bg-hover hover:text-fg"
          >
            引用 {selectionRef}
          </button>
        ) : null}
        {findBar}
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

      {tabs || statusRight ? (
        <div className="flex h-7 shrink-0 items-center gap-2 overflow-x-auto border-t border-line bg-panel px-2">
          {tabs}
          <div className="flex-1" />
          {statusRight}
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
  const readOnly = useAppStore((s) => s.docs.find((d) => d.id === docId)?.readOnly ?? false);
  const [active, setActive] = useState<{ row: number; col: number } | null>(null);
  const [barDraft, setBarDraft] = useState<string | null>(null);
  /** 刚提交、等解析追上的编辑：避免 200ms 防抖期间显示回旧值 */
  const [overlay, setOverlay] = useState<Map<string, string>>(new Map());
  /** 自动换行 / 自动调整列宽行高（与表格路径一致） */
  const [wrapText, setWrapText] = useState(false);
  const [autoFitToken, setAutoFitToken] = useState(0);
  const [autoFitRowsToken, setAutoFitRowsToken] = useState(0);
  const barApiRef = useRef<{ insertReference: (reference: string) => void } | null>(null);
  const handleBarApi = useCallback(
    (api: { insertReference: (reference: string) => void } | null) => {
      barApiRef.current = api;
    },
    [],
  );
  const [barFormulaMode, setBarFormulaMode] = useState(false);
  const handlePickModeChange = useCallback((active: boolean) => setBarFormulaMode(active), []);
  /** 历史深度变化时刷新按钮可用状态（历史本体按文档存在 utils/sheetHistory 里） */
  const [historyVersion, setHistoryVersion] = useState(0);
  const historyDepth = useMemo(() => sheetHistoryDepth(docId), [docId, historyVersion]);
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
  const handlePickReference = useCallback(
    (range: { startRow: number; startCol: number; endRow: number; endCol: number }) => {
      barApiRef.current?.insertReference(
        selectionReference(range, { usedRows: table.rows, usedCols: table.cols }),
      );
    },
    [table],
  );
  const getCell = useCallback(
    (row: number, col: number): SheetCell | undefined => {
      const base = table.rowAt(row)?.[col];
      const pending = overlay.get(`${row},${col}`);
      if (pending === undefined) return base;
      return { v: pending, t: base?.t ?? "text" };
    },
    [overlay, table],
  );

  /* 解析追上内容后清空覆盖层（编辑已经落到文本里了） */
  useEffect(() => {
    setOverlay((prev) => (prev.size === 0 ? prev : new Map()));
  }, [table]);

  /**
   * 把若干单元格的文本写回 CSV（内容立即更新，保存继续走原有文本管线：编码/EOL 都复用）。
   * 一次调用支持多格（清空选区 / 粘贴），底层按行分组只替换改动过的行。
   */
  const writeCells = useCallback(
    (incoming: Array<{ row: number; col: number; value: string }>): boolean => {
      if (incoming.length === 0) return false;
      const latestDoc = useAppStore.getState().docs.find((d) => d.id === docId);
      const next = applyCellEditsToDelimited(latestDoc?.content ?? "", incoming, {
        ext: extName(filePath ?? ""),
      });
      if (next === null) return false;
      useAppStore.getState().patchDoc(docId, {
        content: next,
        // 改回原值时脏标记要能自动消失
        isDirty: next !== (latestDoc?.savedContent ?? ""),
      });
      setOverlay((prev) => {
        const map = new Map(prev);
        for (const item of incoming) map.set(`${item.row},${item.col}`, item.value);
        return map;
      });
      return true;
    },
    [docId, filePath],
  );

  /** 网格提交一格（表格内编辑） */
  const commitCell = useCallback(
    (row: number, col: number, text: string) => {
      if (readOnly) return;
      const before = table.rowAt(row)?.[col]?.v ?? "";
      if (before === text) return;
      if (!writeCells([{ row, col, value: text }])) return;
      pushCellStep<string>(docId, [{ key: `${row},${col}`, before, after: text }]);
      setHistoryVersion((value) => value + 1);
      setBarDraft(null);
    },
    [docId, readOnly, table, writeCells],
  );

  /* ---------------- 选区：统计 / 复制 / 清空 / 粘贴 ---------------- */

  /** Ctrl+C：CSV 数据全在内存里，直接序列化即可 */
  const copySelection = useCallback(
    async (range: SheetRange) => {
      try {
        await navigator.clipboard.writeText(rangeToTsv(range, getCell));
      } catch (error) {
        await showMessage("复制失败", String(error));
      }
    },
    [getCell],
  );

  /** Delete：清空选区 */
  const clearSelection = useCallback(
    async (range: SheetRange) => {
      if (readOnly) return;
      const edits: Array<{ row: number; col: number; before: string; after: string }> = [];
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let col = range.startCol; col <= range.endCol; col += 1) {
          const before = table.rowAt(row)?.[col]?.v ?? "";
          if (before === "") continue;
          edits.push({ row, col, before, after: "" });
        }
      }
      if (edits.length === 0) return;
      if (edits.length > BULK_LIMIT) {
        const ok = await askConfirm({
          title: "清空大片区域",
          message: `将清空 ${edits.length} 个单元格，确定继续吗？`,
          confirmText: "清空",
        });
        if (!ok) return;
      }
      if (!writeCells(edits.map(({ row, col, after }) => ({ row, col, value: after })))) return;
      pushCellStep<string>(
        docId,
        edits.map((item) => ({
          key: `${item.row},${item.col}`,
          before: item.before,
          after: item.after,
        })),
      );
      setHistoryVersion((value) => value + 1);
      setSummary(summarizeRange(range, getCell));
    },
    [docId, readOnly, table, writeCells, getCell],
  );

  /** Ctrl+V：把剪贴板表格按选区左上角写入 */
  const pasteInto = useCallback(
    async (range: SheetRange, text: string) => {
      if (readOnly) return;
      const parsed = parseClipboardTable(text);
      const total = parsed.reduce((sum, row) => sum + row.length, 0);
      if (total === 0) return;
      if (total > BULK_LIMIT) {
        const ok = await askConfirm({
          title: "粘贴大片数据",
          message: `剪贴板包含 ${parsed.length} 行、共 ${total} 个单元格，确定写入吗？`,
          confirmText: "粘贴",
        });
        if (!ok) return;
      }
      const steps: Array<{ row: number; col: number; before: string; after: string }> = [];
      parsed.forEach((values, rowOffset) => {
        values.forEach((value, colOffset) => {
          const row = range.startRow + rowOffset;
          const col = range.startCol + colOffset;
          const before = table.rowAt(row)?.[col]?.v ?? "";
          if (before === value) return;
          steps.push({ row, col, before, after: value });
        });
      });
      if (steps.length === 0) return;
      if (!writeCells(steps.map(({ row, col, after }) => ({ row, col, value: after })))) return;
      pushCellStep<string>(
        docId,
        steps.map((item) => ({
          key: `${item.row},${item.col}`,
          before: item.before,
          after: item.after,
        })),
      );
      setHistoryVersion((value) => value + 1);
      setSummary(summarizeRange(range, getCell));
    },
    [docId, readOnly, table, writeCells, getCell],
  );

  /** 撤销 / 重做一步：历史按文档保存，切标签、切视图（表格/源码/分屏）后依然有效 */
  const applyHistoryBatch = useCallback(
    (batch: CellStep<string>, useBefore: boolean): boolean => {
      const edits = batch.cells.map((cell) => {
        const [rowText, colText] = cell.key.split(",");
        return {
          row: Number(rowText),
          col: Number(colText),
          value: (useBefore ? cell.before : cell.after) ?? "",
        };
      });
      return writeCells(edits);
    },
    [writeCells],
  );

  const stepEdit = useCallback(
    (direction: "undo" | "redo"): boolean => {
      const step = popHistoryStep<string>(docId, direction);
      if (!step) return false;
      // CSV 没有结构操作，历史里只会是单元格步骤
      if (step.kind !== "cells") {
        commitHistoryStep<string>(docId, direction, step);
        return false;
      }
      const ok = applyHistoryBatch(step, direction === "undo");
      if (!ok) {
        returnHistoryStep<string>(docId, direction, step);
        return false;
      }
      commitHistoryStep<string>(docId, direction, step);
      setHistoryVersion((value) => value + 1);
      return true;
    },
    [applyHistoryBatch, docId],
  );

  const undoOnce = useCallback((): boolean => {
    if (sheetHistoryDepth(docId).undo === 0) return false;
    return stepEdit("undo");
  }, [docId, stepEdit]);

  const redoOnce = useCallback((): boolean => {
    if (sheetHistoryDepth(docId).redo === 0) return false;
    return stepEdit("redo");
  }, [docId, stepEdit]);

  const [selection, setSelection] = useState<SheetRange | null>(null);
  const [summary, setSummary] = useState<RangeSummary | null>(null);
  /** 单元格右键菜单（CSV 没有行列结构操作，只提供复制 / 粘贴 / 清空 / 撤销） */
  const [cellMenu, setCellMenu] = useState<{ x: number; y: number; row: number; col: number } | null>(
    null,
  );
  const menuRange = useCallback((): SheetRange | null => {
    if (!cellMenu) return null;
    return (
      selection ?? {
        startRow: cellMenu.row,
        startCol: cellMenu.col,
        endRow: cellMenu.row,
        endCol: cellMenu.col,
      }
    );
  }, [cellMenu, selection]);
  const pasteFromMenu = useCallback(async () => {
    const range = menuRange();
    if (!range) return;
    try {
      const text = await navigator.clipboard.readText();
      await pasteInto(range, text);
    } catch (error) {
      await showMessage("粘贴失败", String(error));
    }
  }, [menuRange, pasteInto]);
  const { columnWidths, rowHeights, handleColumnResize, handleRowResize, handleRowsResize } = useSheetLayout(
    docId,
    () => setHistoryVersion((value) => value + 1),
  );
  /** 冻结的首行/首列（工具栏「冻结」下拉设置，存在文档状态里） */
  const freezeRows = useAppStore((s) => s.docs.find((d) => d.id === docId)?.sheetFreezeRows) ?? 0;
  const freezeCols = useAppStore((s) => s.docs.find((d) => d.id === docId)?.sheetFreezeCols) ?? 0;

  const commitSelection = useCallback(
    (range: SheetRange | null) => {
      setSelection(range);
      setSummary(range ? summarizeRange(range, getCell) : null);
    },
    [getCell],
  );

  const containerRef = useRef<HTMLDivElement | null>(null);
  const isActiveDoc = useAppStore((s) => s.activeId === docId);

  /* ---------------- 查找替换：CSV 数据在内存里，直接按行扫描 ---------------- */

  const findHits = useCallback(
    async (query: string, options: SheetFindOptions, scope: SheetFindScope) => {
      const matcher = buildCellMatcher(query, options);
      if (!matcher) return { hits: [], capped: false, error: "正则表达式无效" };
      const cap = 500;
      // 超大 CSV 只扫前 5 万行，避免长时间阻塞界面（并如实标注被截断）
      const scanRows = Math.min(table.rows, 50000);
      const range = scope === "selection" ? selection : null;
      const hits: SheetFindHit[] = [];
      for (let row = 0; row < scanRows && hits.length < cap; row += 1) {
        const cells = table.rowAt(row);
        if (!cells) continue;
        for (let col = 0; col < cells.length; col += 1) {
          if (!inSelectionRange(range, row, col)) continue;
          if (matcher(cells[col].v)) {
            hits.push({ row, col });
            if (hits.length >= cap) break;
          }
        }
      }
      return { hits, capped: hits.length >= cap || table.rows > scanRows };
    },
    [selection, table],
  );

  const jumpToHit = useCallback((hit: SheetFindHit) => {
    setActive({ row: hit.row, col: hit.col });
    setSelection({ startRow: hit.row, startCol: hit.col, endRow: hit.row, endCol: hit.col });
    setSummary(null);
  }, []);

  /**
   * 替换：把命中格的文本改掉，**整批只记一个撤销步**（CSV 走文本管线，内容直接更新）。
   * `all = true` 时替换每格内的全部出现（对应面板的「全部替换」）。
   */
  const replaceHits = useCallback(
    async (replacement: string, ctx: SheetFindContext, all: boolean): Promise<number> => {
      if (readOnly) return 0;
      const targets = all ? ctx.hits : ctx.hit ? [ctx.hit] : [];
      const edits: Array<{ row: number; col: number; before: string; after: string }> = [];
      for (const hit of targets) {
        const before = table.rowAt(hit.row)?.[hit.col]?.v ?? "";
        const after = replaceCellText(before, lastFindRef.current?.query ?? "", replacement, ctx.options, all);
        if (after === null) continue;
        edits.push({ row: hit.row, col: hit.col, before, after });
      }
      if (edits.length === 0) return 0;
      if (!writeCells(edits.map(({ row, col, after }) => ({ row, col, value: after })))) return 0;
      pushCellStep<string>(
        docId,
        edits.map((item) => ({
          key: `${item.row},${item.col}`,
          before: item.before,
          after: item.after,
        })),
      );
      setHistoryVersion((value) => value + 1);
      return edits.length;
    },
    [docId, readOnly, table, writeCells],
  );

  /** 最近一次查找用的词与开关：面板的替换回调只带 ctx，这里记住产生命中的那一次 */
  const lastFindRef = useRef<{ query: string; options: SheetFindOptions } | null>(null);
  const findApi = useSheetFind({
    isActiveDoc,
    supportsRegex: true,
    find: async (query, options, scope) => {
      lastFindRef.current = { query, options };
      return findHits(query, options, scope);
    },
    jump: jumpToHit,
    replaceCurrent: (replacement, ctx) => replaceHits(replacement, ctx, false),
    replaceAll: (replacement, ctx) => replaceHits(replacement, ctx, true),
  });

  /* 菜单栏命令：一键自动调整表格（CSV 没有公式，不响应插入公式） */
  useEffect(() => {
    if (!isActiveDoc) return;
    return onSheetCommand((command) => {
      if (command.kind === "autoFit") {
        const target = command.target ?? "both";
        if (target !== "rows") setAutoFitToken((value) => value + 1);
        if (target !== "columns") setAutoFitRowsToken((value) => value + 1);
      }
    });
  }, [isActiveDoc]);

  useSheetUndoRedo({
    enabled: !readOnly,
    isActiveDoc,
    containerRef,
    undo: undoOnce,
    redo: redoOnce,
  });

  const activeCell = active && active.row < table.rows ? active : null;
  const activePending = activeCell ? overlay.get(`${activeCell.row},${activeCell.col}`) : undefined;
  const activeValue =
    activePending ?? (activeCell ? (table.rowAt(activeCell.row)?.[activeCell.col]?.v ?? "") : "");
  const barValue = barDraft ?? activeValue;
  const delimiterLabel = table.delimiter === "\t" ? "制表符" : table.delimiter;
  const hint = `${table.rows} 行${table.truncatedCols ? " · 列已截断" : ""}${
    readOnly ? "" : " · 双击可编辑"
  }`;

  /* 切换单元格时清掉公式栏草稿 */
  useEffect(() => {
    setBarDraft(null);
  }, [activeCell?.row, activeCell?.col]);

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
    <div ref={containerRef} className="flex h-full min-h-0 flex-col">
    <SheetFrame
      reference={activeCell ? cellRef(activeCell.row, activeCell.col) : ""}
      content={activeValue}
      usedRows={table.rows}
      usedCols={table.cols}
      onBarApi={handleBarApi}
      onPickModeChange={handlePickModeChange}
      selectionRef={
        barDraft !== null
          ? selection
            ? selectionReference(selection, { usedRows: table.rows, usedCols: table.cols })
            : activeCell
              ? cellRef(activeCell.row, activeCell.col)
              : undefined
          : undefined
      }
      findBar={<SheetFindBar api={findApi} supportsRegex={true} />}
      readOnly={readOnly}
      hint={hint}
      bar={
        readOnly
          ? undefined
          : {
              value: barValue,
              onChange: setBarDraft,
              editing: barDraft !== null,
              onStartFormula: () => setBarDraft("="),
              onCommit: () => {
                if (!activeCell || barDraft === null) return;
                commitCell(activeCell.row, activeCell.col, barDraft);
              },
              onCancel: () => setBarDraft(null),
            }
      }
      tabs={
        <span className="text-[11px] text-faint" title="自动识别的分隔符">
          分隔符：{delimiterLabel} · 列 {table.cols}
        </span>
      }
      statusRight={
        <span className="flex shrink-0 items-center gap-2">
          <button
            type="button"
            onClick={undoOnce}
            disabled={historyDepth.undo === 0}
            title="撤销上一次改动 (Ctrl+Z)"
            className="rounded px-1.5 py-0.5 text-[11px] text-muted hover:bg-hover hover:text-fg disabled:opacity-40"
          >
            撤销{historyDepth.undo > 0 ? ` ${historyDepth.undo}` : ""}
          </button>
          <button
            type="button"
            onClick={redoOnce}
            disabled={historyDepth.redo === 0}
            title="重做 (Ctrl+Y)"
            className="rounded px-1.5 py-0.5 text-[11px] text-muted hover:bg-hover hover:text-fg disabled:opacity-40"
          >
            重做
          </button>
          <button
            type="button"
            onClick={findApi.open}
            title="在表格中查找 (Ctrl+F)"
            className="rounded px-1.5 py-0.5 text-[11px] text-muted hover:bg-hover hover:text-fg"
          >
            查找
          </button>
          <SelectionSummary range={selection} summary={summary} />
        </span>
      }
    >
      <SheetGrid
        rows={table.rows}
        cols={table.cols}
        getCell={getCell}
        /* CSV 首行通常是表头：冻结它，长表滚动时仍能看到列名 */
        headerRow
        editable={!readOnly}
        pendingCells={overlay}
        onCellCommit={commitCell}
        selection={selection}
        onSelectionChange={setSelection}
        onSelectionCommit={commitSelection}
        onClearSelection={(range) => void clearSelection(range)}
        onCopySelection={(range) => void copySelection(range)}
        onPaste={(range, text) => void pasteInto(range, text)}
        onCellContextMenu={(target, position) =>
          setCellMenu({ x: position.x, y: position.y, row: target.row, col: target.col })
        }
        completeFormula={(text, caret) =>
          matchFormulaFunctions(text, caret, { usedRows: table.rows, usedCols: table.cols })
        }
        wrapText={wrapText}
        autoFitToken={autoFitToken}
        autoFitRowsToken={autoFitRowsToken}
        onRowsResize={handleRowsResize}
        pickReference={barFormulaMode}
        onPickReference={handlePickReference}
        formulaRanges={[]}
        columnWidths={columnWidths}
        rowHeights={rowHeights}
        onColumnResize={handleColumnResize}
        onRowResize={handleRowResize}
        freezeRows={freezeRows}
        freezeCols={freezeCols}
        activeCell={activeCell}
        onActiveCell={(row, col) => setActive({ row, col })}
      />
    </SheetFrame>
    {cellMenu ? (
      <ContextMenu
        x={cellMenu.x}
        y={cellMenu.y}
        onClose={() => setCellMenu(null)}
        groups={[
          [
            {
              label: "复制",
              hint: "Ctrl+C",
              onClick: () => {
                const range = menuRange();
                if (range) void copySelection(range);
              },
            },
            { label: "粘贴", hint: "Ctrl+V", disabled: readOnly, onClick: () => void pasteFromMenu() },
            {
              label: "清空内容",
              hint: "Delete",
              disabled: readOnly,
              onClick: () => {
                const range = menuRange();
                if (range) void clearSelection(range);
              },
            },
          ],
          [
            {
              label: "自动调整列宽",
              onClick: () => setAutoFitToken((value) => value + 1),
            },
            {
              label: "自动调整行高",
              onClick: () => setAutoFitRowsToken((value) => value + 1),
            },
            {
              label: wrapText ? "取消自动换行" : "自动换行",
              onClick: () => setWrapText((value) => !value),
            },
          ],
          [
            {
              label: "撤销",
              hint: "Ctrl+Z",
              disabled: historyDepth.undo === 0,
              onClick: () => void undoOnce(),
            },
            {
              label: "重做",
              hint: "Ctrl+Y",
              disabled: historyDepth.redo === 0,
              onClick: () => void redoOnce(),
            },
          ],
        ]}
      />
    ) : null}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* xlsx / xls / ods：Rust 侧解析，按行窗口按需加载（xlsx 支持轻量编辑） */
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
  /** 每做一次结构操作 +1：结构只改内存影子，不会改 modifiedAt，用它触发重读 */
  const [structureVersion, setStructureVersion] = useState(0);
  /** 自动换行（右键菜单切换）：按列宽折行并自动调整行高 */
  const [wrapText, setWrapText] = useState(false);
  /** 自动调整列宽的令牌：值变化时网格按已加载内容重算一次列宽 */
  const [autoFitToken, setAutoFitToken] = useState(0);
  /** 自动调整行高的令牌：值变化时网格把已加载行按内容重算一次行高 */
  const [autoFitRowsToken, setAutoFitRowsToken] = useState(0);
  /** 内容栏「写入区域引用」的 API（网格拖选范围后回填公式） */
  const barApiRef = useRef<{ insertReference: (reference: string) => void } | null>(null);
  const handleBarApi = useCallback(
    (api: { insertReference: (reference: string) => void } | null) => {
      barApiRef.current = api;
    },
    [],
  );
  /** 内容栏是否正在编辑公式：决定网格是否开启"鼠标拾取范围" */
  const [barFormulaMode, setBarFormulaMode] = useState(false);
  const handlePickModeChange = useCallback((active: boolean) => setBarFormulaMode(active), []);
  const { columnWidths, rowHeights, handleColumnResize, handleRowResize, handleRowsResize } = useSheetLayout(
    docId,
    () => setHistoryVersion((value) => value + 1),
  );
  /** 冻结的首行/首列（工具栏「冻结」下拉设置，存在文档状态里） */
  const freezeRows = useAppStore((s) => s.docs.find((d) => d.id === docId)?.sheetFreezeRows) ?? 0;
  const freezeCols = useAppStore((s) => s.docs.find((d) => d.id === docId)?.sheetFreezeCols) ?? 0;

  /** 网格里拖选出的范围 → 写进内容栏的公式 */
  const handlePickReference = useCallback(
    (range: { startRow: number; startCol: number; endRow: number; endCol: number }) => {
      barApiRef.current?.insertReference(
        selectionReference(range, { usedRows: dim?.rows ?? 0, usedCols: dim?.cols ?? 0 }),
      );
    },
    [dim],
  );

  /** 是否可编辑：xlsx 且非只读（xlsm/xls/xlsb/ods 与超大文件只读） */
  const editable = useAppStore((s) => {
    const target = s.docs.find((d) => d.id === docId);
    return Boolean(target && !target.readOnly);
  });
  /** 待提交编辑的唯一数据源放在文档状态里：标签脏标记、关闭确认、会话恢复都靠它 */
  const sheetEdits = useAppStore((s) => s.docs.find((d) => d.id === docId)?.sheetEdits) ?? EMPTY_EDITS;
  /** 历史深度变化时刷新按钮可用状态（历史本体按文档存在 utils/sheetHistory 里） */
  const [historyVersion, setHistoryVersion] = useState(0);
  const historyDepth = useMemo(() => sheetHistoryDepth(docId), [docId, historyVersion]);
  /** 公式栏草稿：null 表示显示当前单元格的值（未在编辑） */
  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);

  const sheetName = info?.sheets[sheetIndex]?.name ?? null;


  /** 公式计算结果缓存：`工作表|行|列|公式文本` → 显示值/错误。仅用于界面显示，不写进文件 */
  const evalCacheRef = useRef<Map<string, { value?: string; error?: string }>>(new Map());
  const [evalVersion, setEvalVersion] = useState(0);

  const evalKeyOf = (sheet: string, row: number, col: number, formula: string) =>
    `${sheet}|${row}|${col}|${formula}`;

  /**
   * 求值一批公式（未落盘的编辑也会参与计算）。
   * 结果只影响显示：文件里仍然只有公式文本，Excel / WPS 打开时会自己重算。
   */
  const evaluate = useCallback(
    async (items: Array<{ row: number; col: number; formula: string }>) => {
      if (!filePath || !sheetName || items.length === 0) return;
      const fresh = items.filter(
        (item) => !evalCacheRef.current.has(evalKeyOf(sheetName, item.row, item.col, item.formula)),
      );
      if (fresh.length === 0) return;
      const overrides = useAppStore.getState().docs.find((d) => d.id === docId)?.sheetEdits ?? [];
      try {
        const results = await invoke<EvalResult[]>("spreadsheet_eval", {
          path: filePath,
          sheet: sheetName,
          requests: fresh.map((item) => ({ row: item.row, col: item.col, formula: item.formula })),
          overrides,
        });
        for (const result of results) {
          const item = fresh.find((entry) => entry.row === result.row && entry.col === result.col);
          if (!item) continue;
          evalCacheRef.current.set(evalKeyOf(sheetName, item.row, item.col, item.formula), {
            value: result.value ?? undefined,
            error: result.error ?? undefined,
          });
        }
        setEvalVersion((value) => value + 1);
      } catch {
        /* 求值失败：单元格继续显示公式原文，不打断编辑 */
      }
    },
    [docId, filePath, sheetName],
  );

  /** 已加载窗口里的公式单元格：交给后端算一遍，免得显示 0（我们写回时不写缓存值） */
  const evaluateLoadedFormulas = useCallback(
    (rows: Array<{ row: number; cells: SheetCell[] }>) => {
      if (!sheetName) return;
      const items: Array<{ row: number; col: number; formula: string }> = [];
      for (const entry of rows) {
        entry.cells.forEach((cell, col) => {
          if (cell?.f) items.push({ row: entry.row, col, formula: cell.f });
        });
      }
      void evaluate(items);
    },
    [evaluate, sheetName],
  );

  /** 当前工作表上待提交的编辑：`行,列` → 文本（网格据此覆盖显示并标记） */
  const pendingCells = useMemo(() => {
    const map = new Map<string, string>();
    if (!sheetName) return map;
    for (const edit of sheetEdits) {
      if (edit.sheet !== sheetName) continue;
      // 公式优先显示后端算出来的结果，算不出来就显示错误，普通编辑显示值本身
      let text = edit.value;
      if (edit.kind === "formula") {
        const cached = evalCacheRef.current.get(evalKeyOf(sheetName, edit.row, edit.col, edit.value));
        if (cached) text = cached.value ?? cached.error ?? edit.value;
      }
      map.set(`${edit.row},${edit.col}`, text);
    }
    return map;
  }, [sheetEdits, sheetName, evalVersion]);

  /**
   * 批量写入编辑（单格提交 / 清空选区 / 粘贴共用）。
   * 一次算出所有 before/after，作为**一个撤销步**压栈 —— 否则清空 20 格要按 20 次 Ctrl+Z。
   */
  const applyEdits = useCallback(
    (incoming: Array<{ row: number; col: number; kind: SheetEditKind; value: string }>) => {
      if (!sheetName || incoming.length === 0) return;
      const list = useAppStore.getState().docs.find((d) => d.id === docId)?.sheetEdits ?? [];
      const steps: EditStep[] = [];
      let merged = list;
      for (const item of incoming) {
        const key = editKey(sheetName, item.row, item.col);
        const pending = merged.find((edit) => editKey(edit.sheet, edit.row, edit.col) === key) ?? null;
        const after: SheetEdit = {
          sheet: sheetName,
          row: item.row,
          col: item.col,
          kind: item.kind,
          value: item.value,
        };
        /**
         * before 必须记下「文件里的原值」而不是 null：
         * 保存之后待提交编辑会被清空，如果 before 只是 null，撤销就退化成「删掉这条待提交编辑」，
         * 界面看到的值仍然是文件里刚保存的新值 —— 表现就是「撤销计数在动，但内容没恢复」。
         * 目标格在文件里本来就是空的时同样要显式写一条 `kind: "empty"` 的编辑
         * （否则用户往空格里输入内容 → 保存 → Ctrl+Z 会完全没有反应）。
         */
        const original = rowsRef.current.get(item.row)?.[item.col];
        const before: SheetEdit =
          pending ??
          (original && (original.v !== "" || original.f)
            ? {
                sheet: sheetName,
                row: item.row,
                col: item.col,
                kind: original.f ? "formula" : inferEditKind(original.v, original.t),
                value: original.f ?? original.v,
              }
            : { sheet: sheetName, row: item.row, col: item.col, kind: "empty", value: "" });
        if (before.kind === after.kind && before.value === after.value) continue;
        steps.push({ key, before, after });
        merged = [...merged.filter((edit) => editKey(edit.sheet, edit.row, edit.col) !== key), after];
      }
      if (steps.length === 0) return;
      pushCellStep<SheetEdit>(docId, steps);
      setHistoryVersion((value) => value + 1);
      useAppStore.getState().patchDoc(docId, { sheetEdits: merged, isDirty: merged.length > 0 });
      // 刚提交的公式立刻算一遍，让界面直接显示结果（而不是公式原文）
      const formulas = steps
        .map((step) => step.after)
        .filter((edit): edit is SheetEdit => edit !== null && edit.kind === "formula")
        .map((edit) => ({ row: edit.row, col: edit.col, formula: edit.value }));
      if (formulas.length > 0) void evaluate(formulas);
    },
    [docId, sheetName, evaluate],
  );

  /**
   * 撤销/重做一整步；历史按文档保存，切标签/切视图/保存之后都依然有效。
   * 步骤分两类：单元格编辑在前端改回值；结构操作交给 Rust 影子的快照栈（spreadsheet_undo/redo）。
   *
   * 用循环而不是单次：后端只为「真的改变了工作簿」的结构操作压快照，空操作（例如在空表上插入行）
   * 在前端栈里可能留下一个没有对应快照的「幽灵步」。遇到就丢掉它继续找下一个，
   * 两条栈自动对齐，不会出现「按了 Ctrl+Z 什么都没发生」。
   */
  const stepEdit = useCallback(
    async (direction: "undo" | "redo"): Promise<boolean> => {
      for (let guard = 0; guard < 20; guard += 1) {
        const step = popHistoryStep<SheetEdit>(docId, direction);
        if (!step) return false;
        if (step.kind === "structure") {
          if (!filePath) {
            returnHistoryStep<SheetEdit>(docId, direction, step);
            return false;
          }
          try {
            const result = await invoke<ShadowEditResult>(
              direction === "undo" ? "spreadsheet_undo" : "spreadsheet_redo",
              { path: filePath },
            );
            // 结构操作被撤销/重做：布局链同步弹出/压回，历史坐标随之回到正确布局
            if (direction === "undo") popLayoutOp(docId, step.op);
            else pushLayoutOpAgain(docId, step.op);
            commitHistoryStep<SheetEdit>(docId, direction, step);
            setHistoryVersion((value) => value + 1);
            // 影子内容变了：刷新表列表、清行缓存重新加载
            setInfo((prev) => (prev ? { ...prev, sheets: result.sheets, pending: result.pending } : prev));
            setSheetIndex((prev) => (prev < result.sheets.length ? prev : 0));
            setStructureVersion((value) => value + 1);
            setSelection(null);
            setSummary(null);
            const pendingEdits =
              useAppStore.getState().docs.find((d) => d.id === docId)?.sheetEdits?.length ?? 0;
            useAppStore.getState().patchDoc(docId, {
              isDirty: result.pending || pendingEdits > 0,
              sheetStructurePending: result.pending,
            });
            return true;
          } catch (error) {
            const message = String(error);
            if (message.includes("没有可撤销") || message.includes("没有可重做")) {
              // 后端没有对应快照 → 这一步是空操作留下的幽灵步：移出栈继续找下一个
              commitHistoryStep<SheetEdit>(docId, direction, step);
              continue;
            }
            returnHistoryStep<SheetEdit>(docId, direction, step);
            await showMessage(direction === "undo" ? "撤销失败" : "重做失败", message);
            return false;
          }
        }

        // 布局步骤（手动列宽/行高）：改回文档状态里的布局覆盖值
        if (step.kind === "layout") {
          const useBefore = direction === "undo";
          const target = useAppStore.getState().docs.find((d) => d.id === docId);
          const cols = { ...(target?.sheetColumnWidths ?? {}) };
          const rows = { ...(target?.sheetRowHeights ?? {}) };
          for (const change of step.changes) {
            // 布局 key 同样按布局版本链换算（结构操作会平移行列，撤销后又要平移回来）
            const mapped = mapCellKey(docId, change.key, step.layoutVersion);
            if (mapped === null) continue;
            const [kind, indexText] = mapped.split(":");
            const index = Number(indexText);
            const value = useBefore ? change.before : change.after;
            const bucket = kind === "col" ? cols : rows;
            if (value === null) delete bucket[index];
            else bucket[index] = value;
          }
          commitHistoryStep<SheetEdit>(docId, direction, step);
          useAppStore.getState().patchDoc(docId, { sheetColumnWidths: cols, sheetRowHeights: rows });
          setHistoryVersion((value) => value + 1);
          return true;
        }

        // 单元格编辑步骤：把 before/after 放回待提交列表
        const useBefore = direction === "undo";
        // 历史坐标是按「产生时的布局」记的：结构操作会平移行列，且结构操作本身也能被撤销，
        // 所以这里统一按布局版本链换算到当前布局（落在已删除行列里的格子会被跳过）
        const cells = mapCellBatch<SheetEdit>(docId, step.cells, step.layoutVersion);
        let list = useAppStore.getState().docs.find((d) => d.id === docId)?.sheetEdits ?? [];
        for (const cell of cells) {
          /**
           * before/after 都是「显式编辑」（空单元格也写成 kind: "empty"），撤销就是把它重新放回待提交列表；
           * 这里的 null 判断只为收窄 utils/sheetHistory 里 CSV 共用的可空类型 —— 命中时说明这一格
           * 本来就等于文件里的值，从待提交列表里移除它就是正确结果。
           */
          const restored = useBefore ? cell.before : cell.after;
          list = list.filter((edit) => editKey(edit.sheet, edit.row, edit.col) !== cell.key);
          if (restored) list = [...list, restored];
        }
        commitHistoryStep<SheetEdit>(docId, direction, step);
        useAppStore.getState().patchDoc(docId, { sheetEdits: list, isDirty: list.length > 0 });
        setHistoryVersion((value) => value + 1);
        return true;
      }
      return false;
    },
    [docId, filePath],
  );

  /** 网格提交一格：推断写回方式后进入待提交列表 */
  const commitCell = useCallback(
    (row: number, col: number, text: string) => {
      if (!editable) return;
      const original = rowsRef.current.get(row)?.[col];
      const kind = inferEditKind(text, original?.t);
      applyEdits([{ row, col, kind, value: kind === "formula" ? text.trim() : text }]);
      setFormulaDraft(null);
    },
    [applyEdits, editable],
  );

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
        // 结构改动只留在内存影子里：用后端返回的 pending 恢复"未保存"提示（例如切换标签后回来）
        if (result.pending) {
          useAppStore.getState().patchDoc(docId, { isDirty: true, sheetStructurePending: true });
        }
        // 保存/外部改动会重新走这里：只在当前工作表不存在时才回到第一张表，避免视野跳走
        setSheetIndex((prev) => (prev < result.sheets.length ? prev : 0));
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

  /* 切换工作表、换文件、文件被改写（保存/外部改动）或结构改动后清空行缓存与选中状态。
     结构操作只改内存影子（modifiedAt 不变），因此额外用 structureVersion 触发重读。 */
  useEffect(() => {
    rowsRef.current.clear();
    pendingRef.current = [];
    setDim(null);
    setActive(null);
    setTruncated(false);
    setVersion(0);
  }, [sheetIndex, filePath, modifiedAt, structureVersion]);

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
              // 本窗口里的公式交给后端算一遍：我们写回时不写缓存值，直接显示会得到 0
              evaluateLoadedFormulas(
                win.cells.map((cells, index) => ({ row: win.start + index, cells })),
              );
            })
            .catch((err) => setError(String(err)))
            .finally(() => {
              pendingRef.current = pendingRef.current.filter(([s, e]) => !(s === from && e === to));
            });
        }
        cursor = requestEnd;
      }
    },
    [filePath, sheetName, evaluateLoadedFormulas],
  );


  // 依赖 version / evalVersion 让 getCell 身份随数据与计算结果变化，网格才会重绘
  const getCell = useCallback(
    (row: number, col: number): SheetCell | undefined => {
      const cached = rowsRef.current.get(row)?.[col];
      // 文件里带公式的单元格：显示后端算出来的结果（写回时我们不写缓存值，读缓存会得到 0）
      if (cached?.f && sheetName) {
        const computed = evalCacheRef.current.get(evalKeyOf(sheetName, row, col, cached.f));
        if (computed) {
          const text = computed.value ?? computed.error ?? cached.v;
          const numeric = computed.value !== undefined && /^-?\d+(\.\d+)?$/.test(computed.value.trim());
          return { ...cached, v: text, t: numeric ? "number" : "text" };
        }
      }
      return cached;
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [version, evalVersion, sheetName],
  );

  /* 首次进入工作表（或缓存被保存/外部改动/结构操作清空）时主动拉第一批数据：
     此时网格还不知道总行数（rows=0），它上报的可视区间为空，无法驱动加载 */
  useEffect(() => {
    if (!sheetName || rowsRef.current.size > 0) return;
    loadRange(0, WINDOW_ROWS);
  }, [sheetName, loadRange, modifiedAt, structureVersion]);

  const onViewport = useCallback(
    // SheetGrid 的 endRow 为开区间，直接透传即可
    (startRow: number, endRow: number) => loadRange(startRow, endRow),
    [loadRange],
  );

  /**
   * 内容栏（公式栏）显示的内容。
   *
   * 关键点：**公式要显示公式原文**，单元格里显示的才是计算结果（和 Excel 一致）。
   * 之前直接用 pendingCells（对公式来说是计算后的值），导致用户输入 `=SUM(B3:B4)` 回车后
   * 公式栏也变成了结果，看起来"公式消失了"。
   */
  const activeFormulaText = useMemo(() => {
    if (!active || !sheetName) return null;
    const pendingEdit = sheetEdits.find(
      (edit) => edit.sheet === sheetName && edit.row === active.row && edit.col === active.col,
    );
    if (pendingEdit?.kind === "formula") return pendingEdit.value;
    return rowsRef.current.get(active.row)?.[active.col]?.f ?? null;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [active, sheetEdits, sheetName, version]);

  /** 当前单元格的显示内容：公式原文 > 待提交编辑（公式为计算结果）> 文件里的值 */
  const activePending = active ? pendingCells.get(`${active.row},${active.col}`) : undefined;
  const activeOriginal = active ? rowsRef.current.get(active.row)?.[active.col] : undefined;
  const activeContent = activeFormulaText ?? activePending ?? activeOriginal?.v ?? "";
  const barValue = formulaDraft ?? activeContent;

  /**
   * 当前单元格公式引用的区域：传给网格画虚线框，显示"这个公式作用在哪一片"。
   * 引用解析支持单格、矩形、整列 B:B、整行 2:2 与带表名的引用（跨表不计入）。
   */
  const formulaRanges = useMemo(
    () =>
      activeFormulaText && sheetName
        ? extractFormulaReferences(activeFormulaText, {
            sheetName,
            usedRows: dim?.rows ?? 0,
            usedCols: dim?.cols ?? 0,
          })
        : [],
    [activeFormulaText, sheetName, dim],
  );

  /** 当前单元格的公式计算错误（显示在内容栏右侧，避免"看起来没反应"） */
  const activeEvalError = useMemo(() => {
    if (!active || !sheetName) return null;
    const pendingEdit = sheetEdits.find(
      (edit) => edit.sheet === sheetName && edit.row === active.row && edit.col === active.col,
    );
    const formula =
      pendingEdit?.kind === "formula"
        ? pendingEdit.value
        : rowsRef.current.get(active.row)?.[active.col]?.f;
    if (!formula) return null;
    return evalCacheRef.current.get(evalKeyOf(sheetName, active.row, active.col, formula))?.error ?? null;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [active, sheetEdits, sheetName, evalVersion]);

  /* 切换单元格/工作表时清掉公式栏草稿 */
  useEffect(() => {
    setFormulaDraft(null);
  }, [active?.row, active?.col, sheetName]);

  /** 撤销 / 重做一步：有步骤就同步返回 true（真正的工作异步进行），供按键决定是否吞掉事件 */
  const undoOnce = useCallback((): boolean => {
    if (sheetHistoryDepth(docId).undo === 0) return false;
    void stepEdit("undo");
    return true;
  }, [docId, stepEdit]);
  const redoOnce = useCallback((): boolean => {
    if (sheetHistoryDepth(docId).redo === 0) return false;
    void stepEdit("redo");
    return true;
  }, [docId, stepEdit]);

  /* ---------------- 选区：统计 / 复制 / 清空 / 粘贴 ---------------- */

  const [selection, setSelection] = useState<SheetRange | null>(null);
  const [summary, setSummary] = useState<RangeSummary | null>(null);

  /** 取值优先看待提交编辑（统计与复制都要反映未保存的改动） */
  const readCell = useCallback<CellReader>(
    (row, col) => {
      const pending = pendingCells.get(`${row},${col}`);
      const base = rowsRef.current.get(row)?.[col];
      if (pending === undefined) return base;
      return { v: pending, t: base?.t ?? "text" };
    },
    [pendingCells],
  );

  /** 选区变化结束后重算统计（拖拽过程中不计算，避免卡顿） */
  const commitSelection = useCallback(
    async (range: SheetRange | null) => {
      setSelection(range);
      if (!range) {
        setSummary(null);
        return;
      }
      const local = summarizeRange(range, readCell);
      setSummary(local);
      // 选区超出已加载窗口时改用 Rust 侧精确统计（不含未保存的编辑，界面不再标 *）
      if (local.partial && filePath && sheetName) {
        try {
          const stats = await invoke<RangeStats>("spreadsheet_stats", {
            path: filePath,
            sheet: sheetName,
            range: {
              startRow: range.startRow,
              startCol: range.startCol,
              endRow: range.endRow,
              endCol: range.endCol,
            },
          });
          setSummary({
            cells: stats.cells,
            nonEmpty: stats.nonEmpty,
            numeric: stats.numeric,
            sum: stats.sum,
            average: stats.average,
            min: stats.min,
            max: stats.max,
            partial: false,
          });
        } catch {
          /* 统计失败时保留本地结果（标注 *部分） */
        }
      }
    },
    [filePath, readCell, sheetName],
  );

  /** Delete：清空选区（写成 empty 编辑，走正常保存流程） */
  const clearSelection = useCallback(
    async (range: SheetRange) => {
      if (!editable || !sheetName) return;
      const count = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
      if (count > BULK_LIMIT) {
        const ok = await askConfirm({
          title: "清空大片区域",
          message: `将清空 ${count} 个单元格，确定继续吗？`,
          confirmText: "清空",
        });
        if (!ok) return;
      }
      const edits: Array<{ row: number; col: number; kind: SheetEditKind; value: string }> = [];
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let col = range.startCol; col <= range.endCol; col += 1) {
          if ((readCell(row, col)?.v ?? "") === "") continue;
          edits.push({ row, col, kind: "empty", value: "" });
        }
      }
      applyEdits(edits);
      setSummary(summarizeRange(range, readCell));
    },
    [applyEdits, editable, readCell, sheetName],
  );

  /** Ctrl+C：复制选区为 TSV（先把选区内的行拉进缓存，保证复制完整） */
  const copySelection = useCallback(
    async (range: SheetRange) => {
      const count = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
      if (count > BULK_LIMIT * 4) {
        await showMessage("选区过大", `一次最多复制 ${BULK_LIMIT * 4} 个单元格，请缩小选区后重试。`);
        return;
      }
      loadRange(range.startRow, range.endRow + 1);
      // 行窗口是异步加载的：给一拍时间再序列化，避免复制出半截数据
      await new Promise((resolve) => window.setTimeout(resolve, 120));
      try {
        await navigator.clipboard.writeText(rangeToTsv(range, readCell));
      } catch (error) {
        await showMessage("复制失败", String(error));
      }
    },
    [loadRange, readCell],
  );

  /** Ctrl+V：把剪贴板表格按选区左上角写入（作为一步撤销） */
  const pasteInto = useCallback(
    async (range: SheetRange, text: string) => {
      if (!editable || !sheetName) return;
      const table = parseClipboardTable(text);
      const cells = table.reduce((sum, row) => sum + row.length, 0);
      if (cells === 0) return;
      if (cells > BULK_LIMIT) {
        const ok = await askConfirm({
          title: "粘贴大片数据",
          message: `剪贴板包含 ${table.length} 行、共 ${cells} 个单元格，确定写入吗？`,
          confirmText: "粘贴",
        });
        if (!ok) return;
      }
      const edits: Array<{ row: number; col: number; kind: SheetEditKind; value: string }> = [];
      table.forEach((values, rowOffset) => {
        values.forEach((value, colOffset) => {
          const row = range.startRow + rowOffset;
          const col = range.startCol + colOffset;
          const original = rowsRef.current.get(row)?.[col];
          edits.push({ row, col, kind: inferEditKind(value, original?.t), value });
        });
      });
      applyEdits(edits);
      setSummary(summarizeRange(range, readCell));
    },
    [applyEdits, editable, readCell, sheetName],
  );

  const containerRef = useRef<HTMLDivElement | null>(null);
  const isActiveDoc = useAppStore((s) => s.activeId === docId);

  /* ---------------- 查找替换：Rust 全表扫描（含企业加密文件） ---------------- */

  const findHits = useCallback(
    async (query: string, options: SheetFindOptions, scope: SheetFindScope) => {
      if (!filePath) return { hits: [], capped: false };
      // 后端只支持区分大小写与全字匹配；正则不支持的要在面板上说明，不能静默当成字面量
      if (options.regex) {
        return {
          hits: [],
          capped: false,
          error: "Excel 表格暂不支持正则查找，可改用「区分大小写 / 全字匹配」",
        };
      }
      const raw = await invoke<
        Array<{ sheet: string; row: number; col: number; text?: string }>
      >("spreadsheet_find", {
        path: filePath,
        query,
        options: { matchCase: options.matchCase, wholeCell: options.wholeCell },
      });
      const range = scope === "selection" ? selection : null;
      const hits = raw
        .filter((hit) => {
          if (scope !== "selection") return true;
          // 选区范围只对当前工作表有意义
          return hit.sheet === sheetName && inSelectionRange(range, hit.row, hit.col);
        })
        .map((hit) => ({ sheet: hit.sheet, row: hit.row, col: hit.col, text: hit.text }));
      return { hits, capped: raw.length >= 500 };
    },
    [filePath, selection, sheetName],
  );

  /**
   * 替换：把命中格的文本改掉并进待提交编辑，**整批只记一个撤销步**。
   * 只作用于**当前工作表**（跨表命中只能查找，面板上有说明）。
   */
  const replaceHits = useCallback(
    async (replacement: string, ctx: SheetFindContext, all: boolean): Promise<number> => {
      if (!editable || !sheetName) return 0;
      const targets = (all ? ctx.hits : ctx.hit ? [ctx.hit] : []).filter(
        (hit) => (hit.sheet ?? sheetName) === sheetName,
      );
      const edits: Array<{ row: number; col: number; kind: SheetEditKind; value: string }> = [];
      for (const hit of targets) {
        const original = rowsRef.current.get(hit.row)?.[hit.col];
        const pendingEdit = sheetEdits.find(
          (edit) => edit.sheet === sheetName && edit.row === hit.row && edit.col === hit.col,
        );
        // 后端把「显示值」和「公式原文」都参与匹配：查询词出现在公式里就替换公式，否则替换显示值
        const formula = pendingEdit?.kind === "formula" ? pendingEdit.value : (original?.f ?? "");
        const displayText = pendingEdit
          ? (pendingEdit.computed ?? pendingEdit.error ?? pendingEdit.value)
          : (original?.v ?? "");
        const query = lastFindRef.current?.query ?? "";
        const matcher = compileQuery(query, ctx.options);
        const source = formula && matcher && matcher.test(formula) ? formula : displayText;
        const next = replaceCellText(source, query, replacement, ctx.options, all);
        if (next === null) continue;
        const kind = inferEditKind(next, original?.t);
        edits.push({
          row: hit.row,
          col: hit.col,
          kind,
          value: kind === "formula" ? next.trim() : next,
        });
      }
      if (edits.length === 0) return 0;
      applyEdits(edits);
      return edits.length;
    },
    [applyEdits, editable, sheetEdits, sheetName],
  );

  const jumpToHit = useCallback(
    (hit: SheetFindHit) => {
      // 跨工作表跳转：先把目标表切过去，等它的数据就绪后再定位（否则会被切换逻辑清掉）
      if (hit.sheet && info) {
        const index = info.sheets.findIndex((sheet) => sheet.name === hit.sheet);
        if (index >= 0 && index !== sheetIndex) {
          pendingJumpRef.current = hit;
          setSheetIndex(index);
          return;
        }
      }
      // 同一张表：先清空再设置，保证网格重新执行「滚动到该格」
      setActive(null);
      requestAnimationFrame(() => {
        setActive({ row: hit.row, col: hit.col });
        setSelection({ startRow: hit.row, startCol: hit.col, endRow: hit.row, endCol: hit.col });
        setSummary(null);
      });
    },
    [info, sheetIndex],
  );

  /* 跨表跳转的收尾：目标表数据就绪后定位到命中格 */
  useEffect(() => {
    const hit = pendingJumpRef.current;
    if (!hit || !sheetName) return;
    if (hit.sheet && hit.sheet !== sheetName) return;
    pendingJumpRef.current = null;
    setActive(null);
    const raf = requestAnimationFrame(() => {
      setActive({ row: hit.row, col: hit.col });
      setSelection({ startRow: hit.row, startCol: hit.col, endRow: hit.row, endCol: hit.col });
    });
    return () => cancelAnimationFrame(raf);
  }, [sheetName, dim]);

  const lastFindRef = useRef<{ query: string; options: SheetFindOptions } | null>(null);
  const findApi = useSheetFind({
    isActiveDoc,
    supportsRegex: false,
    find: async (query, options, scope) => {
      lastFindRef.current = { query, options };
      return findHits(query, options, scope);
    },
    jump: jumpToHit,
    replaceCurrent: (replacement, ctx) => replaceHits(replacement, ctx, false),
    replaceAll: (replacement, ctx) => replaceHits(replacement, ctx, true),
  });

  /* ---------------- 结构操作：插入/删除行列、工作表增删改复制 ---------------- */

  const [sheetMenu, setSheetMenu] = useState<{ x: number; y: number; index: number } | null>(null);
  /** 单元格右键菜单（网格回调） */
  const [cellMenu, setCellMenu] = useState<{
    x: number;
    y: number;
    row: number;
    col: number;
    kind: "cell" | "rowHeader" | "colHeader";
  } | null>(null);
  /** 跨工作表跳转的目标（等目标表数据就绪后再定位） */
  const pendingJumpRef = useRef<SheetFindHit | null>(null);

  /** 网格渲染区域：至少留出可编辑的最小空间（空表也能输入） */
  const gridRows = Math.max(dim?.rows ?? 0, GRID_MIN_ROWS);
  const gridCols = Math.max(dim?.cols ?? 0, GRID_MIN_COLS);

  /** 右键菜单作用的区域：有选区用选区，否则用右键点中的那一格 */
  const menuRange = useCallback((): SheetRange | null => {
    if (!cellMenu) return null;
    return (
      selection ?? {
        startRow: cellMenu.row,
        startCol: cellMenu.col,
        endRow: cellMenu.row,
        endCol: cellMenu.col,
      }
    );
  }, [cellMenu, selection]);

  /** 右键菜单里的粘贴：先读剪贴板 */
  const pasteFromMenu = useCallback(async () => {
    const range = menuRange();
    if (!range) return;
    try {
      const text = await navigator.clipboard.readText();
      await pasteInto(range, text);
    } catch (error) {
      await showMessage("粘贴失败", String(error));
    }
  }, [menuRange, pasteInto]);

  /**
   * 结构操作（插入/删除行列、工作表增删改复制）。
   *
   * 只改 Rust 侧的内存影子工作簿，**不落盘、不弹确认**（和 Excel 一样，删错了按 Ctrl+Z 就能回来）；
   * 按 Ctrl+S 才写回文件。撤销/重做走影子快照栈，保存之后依然可用。
   */
  const runStructure = useCallback(
    async (op: Record<string, unknown>, title: string) => {
      if (!filePath || !editable) return;
      const pending = useAppStore.getState().docs.find((d) => d.id === docId)?.sheetEdits ?? [];
      try {
        // 结构操作只改内存影子工作簿（不落盘），因此这里**不标记自写**
        const result = await invoke<StructureResult>("spreadsheet_structure", {
          path: filePath,
          op,
          edits: pending,
          backup: false,
        });
        useAppStore.getState().patchDoc(docId, {
          // 待提交编辑已经被应用进影子，前端不再单独持有
          sheetEdits: [],
          // pending 由后端按事实给出：越界/空表的空操作不建影子，此时不该把文档标脏
          isDirty: result.pending,
          sheetStructurePending: result.pending,
        });
        // 把这次结构操作记进同一条撤销栈；同时平移已有单元格步骤的坐标，
        // 否则撤销会改到错位的格子（插入行后原来的第 5 行变成第 6 行）
        if (result.pending) {
          // 记进同一条撤销栈，并把「会平移坐标」的结构操作追加到布局版本链；
          // 历史里的单元格坐标在应用时按这条链换算，因此撤销结构操作后旧步骤仍然指向正确的格子
          const kind = String((op as { kind?: unknown }).kind ?? "");
          const layoutOp: LayoutOp | null =
            kind === "insertRows" || kind === "deleteRows" || kind === "insertCols" || kind === "deleteCols"
              ? {
                  kind,
                  sheet: String((op as { sheet?: unknown }).sheet ?? ""),
                  at: Number((op as { at?: unknown }).at ?? 0),
                  count: Number((op as { count?: unknown }).count ?? 1),
                }
              : kind === "renameSheet"
                ? {
                    kind: "renameSheet",
                    sheet: String((op as { sheet?: unknown }).sheet ?? ""),
                    name: String((op as { name?: unknown }).name ?? ""),
                  }
                : kind === "deleteSheet"
                  ? { kind: "deleteSheet", sheet: String((op as { sheet?: unknown }).sheet ?? "") }
                  : null;
          pushStructureStep(docId, title, layoutOp);
        }
        setHistoryVersion((value) => value + 1);
        setSelection(null);
        setSummary(null);
        // 工作表列表可能变了（新建/重命名/删除/复制），且行内容整体平移：重读信息并清空行缓存
        setInfo((prev) => (prev ? { ...prev, sheets: result.sheets, pending: result.pending } : prev));
        setSheetIndex((prev) => (prev < result.sheets.length ? prev : 0));
        setStructureVersion((value) => value + 1);
      } catch (error) {
        await showMessage("操作失败", String(error));
      }
    },
    [docId, editable, filePath],
  );

  const sheetNameAt = useCallback(
    (index: number) => info?.sheets[index]?.name ?? null,
    [info],
  );

  /** 行列插入 / 删除：以当前选区为准（没有选区就用活动单元格，数量为 1） */
  const runRowCol = useCallback(
    async (
      kind: "insertRows" | "deleteRows" | "insertCols" | "deleteCols",
      where: "above" | "below" | "left" | "right" | "current",
    ) => {
      if (!sheetName) return;
      const anchorRow = selection?.startRow ?? active?.row ?? 0;
      const anchorCol = selection?.startCol ?? active?.col ?? 0;
      const rowCount = selection ? selection.endRow - selection.startRow + 1 : 1;
      const colCount = selection ? selection.endCol - selection.startCol + 1 : 1;
      const usedRows = dim?.rows ?? 0;
      const usedCols = dim?.cols ?? 0;
      // 数据区之外的位置：插入/删除都是空操作，给一句提示即可，不必写盘、也不该报错
      if ((kind === "insertRows" || kind === "deleteRows") && anchorRow >= usedRows) {
        await showMessage(
          kind === "insertRows" ? "无需插入" : "没有可删除的行",
          `数据区只到第 ${usedRows} 行，第 ${anchorRow + 1} 行还在数据之外。\n\n直接在网格里输入即可，写回时会把这一行补上。`,
        );
        return;
      }
      if ((kind === "insertCols" || kind === "deleteCols") && anchorCol >= usedCols) {
        await showMessage(
          kind === "insertCols" ? "无需插入" : "没有可删除的列",
          `数据区只到第 ${usedCols} 列，第 ${anchorCol + 1} 列还在数据之外。\n\n直接在网格里输入即可。`,
        );
        return;
      }
      if (kind === "insertRows") {
        const at = where === "below" ? anchorRow + rowCount : anchorRow;
        await runStructure({ kind, sheet: sheetName, at, count: rowCount }, `插入 ${rowCount} 行`);
      } else if (kind === "deleteRows") {
        // 选区可能跨到数据区之外：只删实际存在的部分
        const count = Math.max(1, Math.min(rowCount, usedRows - anchorRow));
        await runStructure({ kind, sheet: sheetName, at: anchorRow, count }, `删除 ${count} 行`);
      } else if (kind === "insertCols") {
        const at = where === "right" ? anchorCol + colCount : anchorCol;
        await runStructure({ kind, sheet: sheetName, at, count: colCount }, `插入 ${colCount} 列`);
      } else {
        const count = Math.max(1, Math.min(colCount, usedCols - anchorCol));
        await runStructure({ kind, sheet: sheetName, at: anchorCol, count }, `删除 ${count} 列`);
      }
    },
    [active, dim, runStructure, selection, sheetName],
  );

  /** 新建 / 重命名 / 删除 / 复制工作表 */
  const runSheetOp = useCallback(
    async (kind: "add" | "rename" | "delete" | "copy", index: number) => {
      const current = sheetNameAt(index);
      // 新建 / 复制会追加到末尾：切过去，否则用户看到的是旧表，容易以为「新建了却没显示」
      const appendedIndex = info?.sheets.length ?? 0;
      if (kind === "add") {
        await runStructure({ kind: "addSheet", name: null }, "新建工作表");
        setSheetIndex(appendedIndex);
        return;
      }
      if (!current) return;
      if (kind === "delete") {
        if (!info || info.sheets.length <= 1) {
          await showMessage("无法删除", "工作簿至少要保留一个工作表。");
          return;
        }
        await runStructure({ kind: "deleteSheet", sheet: current }, `删除工作表「${current}」`);
        return;
      }
      const answer = await askForm({
        title: kind === "rename" ? "重命名工作表" : "复制工作表",
        fields: [
          {
            key: "name",
            label: "名称",
            value: kind === "rename" ? current : `${current} 副本`,
            placeholder: "工作表名称",
            autofocus: true,
          },
        ],
        confirmText: kind === "rename" ? "重命名" : "复制",
      });
      if (!answer) return;
      const name = (answer.name ?? "").trim();
      if (name === "") return;
      if (kind === "rename") {
        if (name === current) return;
        await runStructure({ kind: "renameSheet", sheet: current, name }, `重命名为「${name}」`);
      } else {
        await runStructure({ kind: "copySheet", sheet: current, name }, `复制为「${name}」`);
        setSheetIndex(appendedIndex);
      }
    },
    [info, runStructure, sheetNameAt],
  );

  /* 菜单栏命令：一键自动调整表格 / 插入常用公式（选区作为公式参数） */
  useEffect(() => {
    if (!isActiveDoc) return;
    return onSheetCommand((command) => {
      if (command.kind === "autoFit") {
        const target = command.target ?? "both";
        if (target !== "rows") setAutoFitToken((value) => value + 1);
        if (target !== "columns") setAutoFitRowsToken((value) => value + 1);
        return;
      }
      if (command.kind === "freeze") {
        useAppStore.getState().patchDoc(docId, {
          sheetFreezeRows: command.rows,
          sheetFreezeCols: command.cols,
        });
        return;
      }
      if (command.kind !== "insertFormula" || !editable || !active) return;
      const reference = selection
        ? selectionReference(selection, { usedRows: dim?.rows ?? 0, usedCols: dim?.cols ?? 0 })
        : cellRef(active.row, active.col);
      applyEdits([
        {
          row: active.row,
          col: active.col,
          kind: "formula",
          value: command.template.replace("{range}", reference),
        },
      ]);
    });
  }, [active, applyEdits, dim, editable, isActiveDoc, selection]);

  useSheetUndoRedo({
    enabled: editable,
    isActiveDoc,
    containerRef,
    undo: undoOnce,
    redo: redoOnce,
  });

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
    <div ref={containerRef} className="flex h-full min-h-0 flex-col">
    <SheetFrame
      reference={active ? cellRef(active.row, active.col) : ""}
      content={activeContent}
      usedRows={dim?.rows ?? 0}
      usedCols={dim?.cols ?? 0}
      onBarApi={handleBarApi}
      onPickModeChange={handlePickModeChange}
      selectionRef={
        formulaDraft !== null
          ? selection
            ? selectionReference(selection, { usedRows: dim?.rows ?? 0, usedCols: dim?.cols ?? 0 })
            : active
              ? cellRef(active.row, active.col)
              : undefined
          : undefined
      }
      findBar={<SheetFindBar api={findApi} supportsRegex={true} />}
      readOnly={!editable}
      encrypted={info.encrypted}
      hint={
        activeEvalError ? (
          <span className="text-warning" title="公式计算结果（仅界面显示，文件里仍是公式）">
            ⚠ {activeEvalError}
          </span>
        ) : dim ? (
          `${dim.rows} 行 × ${dim.cols} 列${truncated ? " · 已截断" : ""}${
            sheetEdits.length > 0 ? ` · 已改 ${sheetEdits.length} 格` : ""
          }`
        ) : (
          "正在加载…"
        )
      }
      bar={
        editable
          ? {
              value: barValue,
              onChange: setFormulaDraft,
              editing: formulaDraft !== null,
              onStartFormula: () => setFormulaDraft("="),
              onCommit: () => {
                if (!active || formulaDraft === null) return;
                commitCell(active.row, active.col, formulaDraft);
              },
              onCancel: () => setFormulaDraft(null),
            }
          : undefined
      }
      tabs={
        <>
          {info.sheets.map((sheet, index) => (
            <button
              key={`${sheet.name}-${index}`}
              type="button"
              onClick={() => setSheetIndex(index)}
              onContextMenu={(event) => {
                event.preventDefault();
                if (!editable) return;
                setSheetMenu({ x: event.clientX, y: event.clientY, index });
              }}
              title={editable ? `${sheet.name}（右键管理工作表）` : sheet.name}
              className={`max-w-[180px] shrink-0 truncate rounded px-2 py-0.5 text-[11px] transition-colors ${
                index === sheetIndex
                  ? "bg-accent/15 font-medium text-accent"
                  : "text-muted hover:bg-hover hover:text-fg"
              }`}
            >
              {sheet.name}
            </button>
          ))}
          {editable ? (
            <button
              type="button"
              onClick={() => void runSheetOp("add", sheetIndex)}
              title="新建工作表"
              className="shrink-0 rounded px-1.5 py-0.5 text-[11px] text-muted hover:bg-hover hover:text-fg"
            >
              ＋
            </button>
          ) : null}
        </>
      }
      statusRight={
        <span className="flex shrink-0 items-center gap-2">
          <button
            type="button"
            onClick={undoOnce}
            disabled={historyDepth.undo === 0}
            title="撤销上一次改动 (Ctrl+Z)"
            className="rounded px-1.5 py-0.5 text-[11px] text-muted hover:bg-hover hover:text-fg disabled:opacity-40"
          >
            撤销{historyDepth.undo > 0 ? ` ${historyDepth.undo}` : ""}
          </button>
          <button
            type="button"
            onClick={redoOnce}
            disabled={historyDepth.redo === 0}
            title="重做 (Ctrl+Y)"
            className="rounded px-1.5 py-0.5 text-[11px] text-muted hover:bg-hover hover:text-fg disabled:opacity-40"
          >
            重做
          </button>
          <button
            type="button"
            onClick={findApi.open}
            title="在表格中查找 (Ctrl+F)"
            className="rounded px-1.5 py-0.5 text-[11px] text-muted hover:bg-hover hover:text-fg"
          >
            查找
          </button>
          <SelectionSummary range={selection} summary={summary} />
        </span>
      }
    >
      <SheetGrid
        rows={gridRows}
        cols={gridCols}
        getCell={getCell}
        onViewport={onViewport}
        editable={editable}
        pendingCells={pendingCells}
        onCellCommit={commitCell}
        selection={selection}
        onSelectionChange={setSelection}
        onSelectionCommit={(range) => void commitSelection(range)}
        onClearSelection={(range) => void clearSelection(range)}
        onCopySelection={(range) => void copySelection(range)}
        onPaste={(range, text) => void pasteInto(range, text)}
        onCellContextMenu={(target, position) =>
          setCellMenu({
            x: position.x,
            y: position.y,
            row: target.row,
            col: target.col,
            kind: target.kind,
          })
        }
        completeFormula={(text, caret) =>
          matchFormulaFunctions(text, caret, { usedRows: dim?.rows ?? 0, usedCols: dim?.cols ?? 0 })
        }
        wrapText={wrapText}
        autoFitToken={autoFitToken}
        autoFitRowsToken={autoFitRowsToken}
        onRowsResize={handleRowsResize}
        pickReference={barFormulaMode}
        onPickReference={handlePickReference}
        formulaRanges={formulaRanges}
        columnWidths={columnWidths}
        rowHeights={rowHeights}
        onColumnResize={handleColumnResize}
        onRowResize={handleRowResize}
        freezeRows={freezeRows}
        freezeCols={freezeCols}
        activeCell={active}
        onActiveCell={(row, col) => setActive({ row, col })}
      />
    </SheetFrame>
    {cellMenu ? (
      <ContextMenu
        x={cellMenu.x}
        y={cellMenu.y}
        onClose={() => setCellMenu(null)}
        groups={[
          [
            {
              label: "复制",
              hint: "Ctrl+C",
              onClick: () => {
                const range = menuRange();
                if (range) void copySelection(range);
              },
            },
            { label: "粘贴", hint: "Ctrl+V", disabled: !editable, onClick: () => void pasteFromMenu() },
            {
              label: "清空内容",
              hint: "Delete",
              disabled: !editable,
              onClick: () => {
                const range = menuRange();
                if (range) void clearSelection(range);
              },
            },
          ],
          [
            {
              label: "自动调整列宽",
              onClick: () => setAutoFitToken((value) => value + 1),
            },
            {
              label: "自动调整行高",
              onClick: () => setAutoFitRowsToken((value) => value + 1),
            },
            {
              label: wrapText ? "取消自动换行" : "自动换行",
              onClick: () => setWrapText((value) => !value),
            },
          ],
          [
            {
              label: "撤销",
              hint: "Ctrl+Z",
              disabled: historyDepth.undo === 0,
              onClick: () => void undoOnce(),
            },
            {
              label: "重做",
              hint: "Ctrl+Y",
              disabled: historyDepth.redo === 0,
              onClick: () => void redoOnce(),
            },
          ],
          [
            {
              label: "在上方插入行",
              disabled: !editable,
              onClick: () => void runRowCol("insertRows", "above"),
            },
            {
              label: "在下方插入行",
              disabled: !editable,
              onClick: () => void runRowCol("insertRows", "below"),
            },
            {
              label: cellMenu.kind === "rowHeader" ? "删除该行" : "删除选中行",
              disabled: !editable,
              onClick: () => void runRowCol("deleteRows", "current"),
            },
          ],
          [
            {
              label: "在左侧插入列",
              disabled: !editable,
              onClick: () => void runRowCol("insertCols", "left"),
            },
            {
              label: "在右侧插入列",
              disabled: !editable,
              onClick: () => void runRowCol("insertCols", "right"),
            },
            {
              label: cellMenu.kind === "colHeader" ? "删除该列" : "删除选中列",
              disabled: !editable,
              onClick: () => void runRowCol("deleteCols", "current"),
            },
          ],
        ]}
      />
    ) : null}
    {sheetMenu ? (
      <ContextMenu
        x={sheetMenu.x}
        y={sheetMenu.y}
        onClose={() => setSheetMenu(null)}
        groups={[
          [{ label: "新建工作表", onClick: () => void runSheetOp("add", sheetMenu.index) }],
          [
            { label: "重命名工作表…", onClick: () => void runSheetOp("rename", sheetMenu.index) },
            { label: "复制工作表…", onClick: () => void runSheetOp("copy", sheetMenu.index) },
            {
              label: `删除「${sheetNameAt(sheetMenu.index) ?? ""}」`,
              disabled: (info?.sheets.length ?? 0) <= 1,
              onClick: () => void runSheetOp("delete", sheetMenu.index),
            },
          ],
        ]}
      />
    ) : null}
    </div>
  );
}
