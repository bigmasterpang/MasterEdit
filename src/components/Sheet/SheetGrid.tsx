/**
 * 只读虚拟化表格网格（Excel 查看器风格）。
 * - 行虚拟化：外层滚动容器 + 内层等高占位块（height = 行数 × ROW_H），只渲染与视口相交的行
 *   （上下各加 OVERSCAN 行缓冲），行绝对定位（top = 虚拟行号 × ROW_H），百万行也无压力。
 * - 冻结装饰：列字母表头 sticky top、行号栏 sticky left、左上角单元格两个方向都 sticky。
 * - 列宽：采样前 60 行估算，memo 只依赖「列数 / 总行数 / getCell 引用」，滚动不重采样。
 */
import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { JSX } from "react";
import type { SheetCell, SheetCellType } from "../../types";

/* -------------------------------- 几何常量 -------------------------------- */

/** 行高（虚拟化唯一基准）、表头高度、行号栏宽度、上下/左右 overscan、列宽采样行数与夹取范围 */
const ROW_H = 24;
const HEADER_H = 24;
const GUTTER_W = 56;
/**
 * 上下额外渲染的行数。
 * 取半屏左右（16 行）是为了应对「拖动滚动条快速跳转」：跳转是瞬时的，
 * 缓冲区越大，视口边缘越不容易出现尚未渲染的空白。
 */
const OVERSCAN = 16;
const COL_OVERSCAN = 1;
const SAMPLE_ROWS = 60;
const MIN_COL_W = 56;
const MAX_COL_W = 320;
const DEFAULT_COL_W = 120;

type GetCell = (row: number, col: number) => SheetCell | undefined;

export interface SheetGridProps {
  /** 数据总行数（启用 headerRow 时 row 0 即冻结表头行，数据行号范围 0..rows-1） */
  rows: number;
  /** 数据总列数 */
  cols: number;
  /** 按数据坐标取单元格；返回 undefined 表示该窗口尚未加载（显示为空白） */
  getCell: GetCell;
  /** 可视行范围变化回调（startRow 含、endRow 不含，均为数据行号），用于按需加载窗口 */
  onViewport?: (startRow: number, endRow: number) => void;
  /** 把数据第 0 行当作冻结表头行渲染 */
  headerRow?: boolean;
  /** 显示「只读」徽标 / 「已截断」提示 */
  readOnly?: boolean;
  truncated?: boolean;
  /** 当前活动单元格；传 undefined 表示网格自己管理选中，传 null 表示显式清空 */
  activeCell?: { row: number; col: number } | null;
  /** 活动单元格变化回调 */
  onActiveCell?: (row: number, col: number) => void;
}

/* -------------------------------- 小工具 -------------------------------- */

/** 0 → A，25 → Z，26 → AA（Excel 列名规则） */
function columnLabel(index: number): string {
  let label = "";
  let n = index + 1;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

/** 类型 → 对齐与颜色：数字/日期右对齐并用等宽数字，布尔居中，错误用危险色 */
const CELL_ALIGN: Record<SheetCellType, string> = {
  empty: "text-left",
  text: "text-left",
  number: "text-right tabular-nums",
  date: "text-right tabular-nums",
  bool: "text-center",
  error: "text-left text-danger",
};

/** 粗略「视觉宽度」：CJK / 全角字符按 2 个字符计，用于估算列宽 */
function visualLength(text: string): number {
  let length = 0;
  for (let i = 0; i < text.length; i += 1) length += text.charCodeAt(i) > 0x2e7f ? 2 : 1;
  return length;
}

/**
 * 列宽采样：只看前 sampleRows 行的内容长度（12px 字号约 7.2px/字符，再加左右内边距），
 * 结果夹在 [MIN_COL_W, MAX_COL_W]；采样窗口内没有任何内容的列保持默认宽度。
 */
function computeColWidths(cols: number, sampleRows: number, getCell: GetCell): number[] {
  const widths = new Array<number>(cols).fill(DEFAULT_COL_W);
  if (cols === 0 || sampleRows === 0) return widths;
  const longest = new Array<number>(cols).fill(0);
  for (let row = 0; row < sampleRows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      const cell = getCell(row, col);
      if (!cell || cell.v.length === 0) continue;
      const length = visualLength(cell.v);
      if (length > longest[col]) longest[col] = length;
    }
  }
  for (let col = 0; col < cols; col += 1) {
    if (longest[col] === 0) continue;
    widths[col] = Math.min(MAX_COL_W, Math.max(MIN_COL_W, 18 + Math.round(longest[col] * 7.2)));
  }
  return widths;
}

/* -------------------------------- 单行渲染 -------------------------------- */

interface SheetRowProps {
  /** 数据行号；label 为行号栏文本（1 起） */
  row: number;
  label: string;
  /** 绝对定位的 top（冻结行则是 sticky 的 top）；width 必须等于完整内容宽度，
   *  否则行号栏 sticky 会在横向滚动时“掉队” */
  top: number;
  width: number;
  frozen: boolean;
  firstCol: number;
  lastCol: number;
  colOffsets: number[];
  colWidths: number[];
  getCell: GetCell;
  /** 该行活动单元格所在列，-1 表示该行没有活动单元格 */
  activeCol: number;
  onCellActivate: (row: number, col: number) => void;
}

/**
 * 单行渲染。用 memo 包裹：滚动时只有新进入可视区的行会挂载/重渲染，
 * 活动单元格变化也只影响「失去选中」和「获得选中」的那两行。
 */
const SheetRow = memo(function SheetRow({
  row,
  label,
  top,
  width,
  frozen,
  firstCol,
  lastCol,
  colOffsets,
  colWidths,
  getCell,
  activeCol,
  onCellActivate,
}: SheetRowProps): JSX.Element {
  const cells: JSX.Element[] = [];
  for (let col = firstCol; col <= lastCol; col += 1) {
    const cell = getCell(row, col);
    const text = cell ? cell.v : "";
    const isActive = col === activeCol;
    cells.push(
      // 只画右边和下边：相邻单元格共用一条 1px 线，不会出现双线
      <div
        key={col} role="gridcell" aria-colindex={col + 1} aria-selected={isActive}
        title={text.length > 0 ? text : undefined} onMouseDown={() => onCellActivate(row, col)}
        className={`absolute top-0 truncate border-r border-b border-line px-1.5 text-[12px] text-fg ${
          CELL_ALIGN[cell ? cell.t : "empty"]
        }${isActive ? " z-[1] bg-accent-soft ring-2 ring-inset ring-accent" : ""}${frozen ? " bg-panel font-medium" : ""}`}
        style={{ left: colOffsets[col], width: colWidths[col], height: ROW_H, lineHeight: `${ROW_H}px` }}
      >
        {text}
      </div>,
    );
  }

  const gutterClass = `sticky left-0 border-b border-r border-line text-center text-[11px] ${
    activeCol >= 0 ? "bg-accent-soft text-accent" : "bg-panel text-muted"
  } ${frozen ? "z-20" : "z-10"}`;
  return (
    <div
      role="row"
      className={frozen ? "sticky z-[15] border-b border-line bg-panel" : "absolute left-0"}
      style={{ top, width, height: ROW_H }}
    >
      <div role="rowheader" className={gutterClass} style={{ width: GUTTER_W, height: ROW_H, lineHeight: `${ROW_H}px` }}>
        {label}
      </div>
      {cells}
    </div>
  );
});

/* -------------------------------- 主组件 -------------------------------- */

export function SheetGrid({
  rows,
  cols,
  getCell,
  onViewport,
  headerRow = false,
  readOnly = false,
  truncated = false,
  activeCell,
  onActiveCell,
}: SheetGridProps): JSX.Element {
  const scrollerRef = useRef<HTMLDivElement>(null);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [viewport, setViewport] = useState({ w: 0, h: 0 });
  const [active, setActive] = useState<{ row: number; col: number } | null>(activeCell ?? null);
  const lastViewportKeyRef = useRef("");

  // 外部回调放进 ref：事件回调因此保持引用稳定，父组件重渲染不会导致所有行重新渲染
  const callbacksRef = useRef({ onActiveCell });
  callbacksRef.current = { onActiveCell };

  const empty = rows === 0 || cols === 0;
  const headerOffset = headerRow ? 1 : 0;
  /** 参与滚动的行数：启用表头时第一行被冻结，不再占用滚动高度 */
  const scrollRows = Math.max(0, rows - headerOffset);
  /** 顶部冻结区总高度（列字母表头 + 可选的数据表头行） */
  const stickyH = HEADER_H + headerOffset * ROW_H;

  // 列宽只在列数 / 总行数 / getCell 引用（数据身份）变化时重算，滚动时不重算
  const colWidths = useMemo(() => computeColWidths(cols, Math.min(rows, SAMPLE_ROWS), getCell), [cols, rows, getCell]);

  /** 每列的左偏移（已含行号栏宽度），供绝对定位与滚动定位使用 */
  const colOffsets = useMemo(() => {
    const offsets = new Array<number>(cols);
    let x = GUTTER_W;
    for (let col = 0; col < cols; col += 1) {
      offsets[col] = x;
      x += colWidths[col];
    }
    return offsets;
  }, [cols, colWidths]);

  // 内容宽度至少铺满视口，保证表头背景与右侧空白区域完整
  const contentWidth = Math.max(cols > 0 ? colOffsets[cols - 1] + colWidths[cols - 1] : GUTTER_W, viewport.w);

  /**
   * 滚动位置用 requestAnimationFrame 合并：
   * 拖动滚动条时 scroll 事件可能一帧触发多次，逐次 setState 会让 React 落后于滚动位置，
   * 表现为「先白屏、再补上数据」。合并后每帧最多渲染一次，且总是按最新位置渲染。
   */
  const pendingScrollRef = useRef<{ top: number; left: number } | null>(null);
  const rafRef = useRef<number | null>(null);

  const handleScroll = useCallback(() => {
    const el = scrollerRef.current;
    if (!el) return;
    pendingScrollRef.current = { top: el.scrollTop, left: el.scrollLeft };
    if (rafRef.current !== null) return;
    rafRef.current = window.requestAnimationFrame(() => {
      rafRef.current = null;
      const next = pendingScrollRef.current;
      if (!next) return;
      // 值未变化时返回原对象，React 会跳过这次渲染
      setScroll((prev) => (prev.top === next.top && prev.left === next.left ? prev : next));
    });
  }, []);

  useEffect(
    () => () => {
      if (rafRef.current !== null) window.cancelAnimationFrame(rafRef.current);
    },
    [],
  );

  useLayoutEffect(() => {
    const el = scrollerRef.current;
    if (!el) return;
    const measure = () => {
      const w = el.clientWidth;
      const h = el.clientHeight;
      setViewport((prev) => (prev.w === w && prev.h === h ? prev : { w, h }));
    };
    measure(); // 首帧同步测量，避免“先只渲染一列/一行再补满”的闪烁
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  }, [empty]);

  const viewRange = useMemo(() => {
    if (empty) return { firstRow: 0, lastRow: -1, firstCol: 0, lastCol: -1, firstVisible: 0, lastVisible: -1 };
    // 虚拟化核心：scrollTop / ROW_H 就是第一个可见的虚拟行；
    // 可见行数由「视口高度 - 顶部冻结区高度」决定，最后上下各加 OVERSCAN 行缓冲。
    const firstVisible = Math.floor(scroll.top / ROW_H);
    const lastVisible = Math.max(firstVisible, Math.floor((scroll.top + Math.max(0, viewport.h - stickyH)) / ROW_H));
    const firstRow = Math.max(0, firstVisible - OVERSCAN);
    const lastRow = Math.min(scrollRows - 1, lastVisible + OVERSCAN);

    // 列窗口：横向滚动时左侧有 GUTTER_W 被行号栏遮住，可见区左边界要加上它
    const left = scroll.left + GUTTER_W;
    const right = scroll.left + Math.max(viewport.w, GUTTER_W);
    let firstCol = 0;
    while (firstCol < cols - 1 && colOffsets[firstCol] + colWidths[firstCol] <= left) firstCol += 1;
    let lastCol = cols - 1;
    while (lastCol > 0 && colOffsets[lastCol] >= right) lastCol -= 1;

    return {
      firstRow,
      lastRow,
      firstCol: Math.max(0, firstCol - COL_OVERSCAN),
      lastCol: Math.min(cols - 1, lastCol + COL_OVERSCAN),
      firstVisible,
      lastVisible,
    };
  }, [empty, scroll.top, scroll.left, viewport.w, viewport.h, stickyH, scrollRows, cols, colOffsets, colWidths]);

  const visibleRows: number[] = [];
  const visibleCols: number[] = [];
  for (let v = viewRange.firstRow; v <= viewRange.lastRow; v += 1) visibleRows.push(v + headerOffset);
  for (let col = viewRange.firstCol; col <= viewRange.lastCol; col += 1) visibleCols.push(col);

  // 可视行范围上报：key 带上行列总数，换表时即使范围相同也会重新上报；
  // 同一范围只上报一次，避免滚动过程中反复触发按需加载。
  // rows/cols 仍为 0（例如工作簿首帧还没拿到维度）时也要在挂载时上报一次 onViewport(0, 0)，
  // 否则纯按需加载的父组件永远等不到第一次请求，会一直卡在空表状态。
  useEffect(() => {
    const startRow = Math.min(rows, viewRange.firstVisible + headerOffset);
    const endRow = Math.min(rows, Math.max(viewRange.firstVisible + 1, viewRange.lastVisible + 1 + headerOffset));
    const key = `${rows}:${cols}:${startRow}:${endRow}`;
    if (lastViewportKeyRef.current === key) return;
    lastViewportKeyRef.current = key;
    onViewport?.(startRow, endRow);
  }, [rows, cols, headerOffset, onViewport, viewRange.firstVisible, viewRange.lastVisible]);

  // 受控同步：undefined 表示非受控（网格自己管选中），null 表示显式清空；
  // 值相同时复用原对象，避免父组件每次传新字面量导致多余渲染
  useEffect(() => {
    if (activeCell === undefined) return;
    if (activeCell === null) {
      setActive(null); // 已是 null 时 React 会跳过这次渲染
      return;
    }
    if (rows === 0 || cols === 0) return;
    const row = Math.max(0, Math.min(rows - 1, activeCell.row));
    const col = Math.max(0, Math.min(cols - 1, activeCell.col));
    setActive((prev) => (prev && prev.row === row && prev.col === col ? prev : { row, col }));
  }, [activeCell, rows, cols]);

  /** 把单元格滚入可视区（纵向避开顶部冻结区，横向避开左侧行号栏） */
  const ensureVisible = useCallback(
    (row: number, col: number) => {
      const el = scrollerRef.current;
      if (!el) return;
      const y = stickyH + (row - headerOffset) * ROW_H;
      if (y < el.scrollTop + stickyH) el.scrollTop = Math.max(0, y - stickyH);
      else if (y + ROW_H > el.scrollTop + el.clientHeight) el.scrollTop = y + ROW_H - el.clientHeight;

      const x = colOffsets[col] ?? GUTTER_W;
      const width = colWidths[col] ?? DEFAULT_COL_W;
      if (x < el.scrollLeft + GUTTER_W) el.scrollLeft = Math.max(0, x - GUTTER_W);
      else if (x + width > el.scrollLeft + el.clientWidth) el.scrollLeft = x + width - el.clientWidth;
    },
    [stickyH, headerOffset, colOffsets, colWidths],
  );

  const handleCellActivate = useCallback(
    (row: number, col: number) => {
      setActive({ row, col });
      callbacksRef.current.onActiveCell?.(row, col);
      scrollerRef.current?.focus({ preventScroll: true }); // 点击后让网格拿到键盘焦点
      ensureVisible(row, col);
    },
    [ensureVisible],
  );

  const moveActive = useCallback(
    (row: number, col: number) => {
      const nextRow = Math.max(0, Math.min(rows - 1, row));
      const nextCol = Math.max(0, Math.min(cols - 1, col));
      setActive({ row: nextRow, col: nextCol });
      callbacksRef.current.onActiveCell?.(nextRow, nextCol);
      ensureVisible(nextRow, nextCol);
    },
    [rows, cols, ensureVisible],
  );

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (empty) return;
      const current = active ?? { row: 0, col: 0 };

      if ((event.ctrlKey || event.metaKey) && (event.key === "c" || event.key === "C")) {
        // 复制当前单元格的显示文本；剪贴板不可用（非安全上下文等）时静默忽略
        try {
          void navigator.clipboard.writeText(getCell(current.row, current.col)?.v ?? "").catch(() => {});
        } catch {
          /* 忽略剪贴板异常 */
        }
        event.preventDefault();
        return;
      }

      // 翻页步长：一屏可见行数再留 1 行重叠，与 Excel 手感一致
      const pageSize = Math.max(1, Math.floor(Math.max(0, viewport.h - stickyH) / ROW_H) - 1);
      switch (event.key) {
        case "ArrowUp": moveActive(current.row - 1, current.col); break;
        case "ArrowDown": moveActive(current.row + 1, current.col); break;
        case "ArrowLeft": moveActive(current.row, current.col - 1); break;
        case "ArrowRight": moveActive(current.row, current.col + 1); break;
        case "Home": moveActive(current.row, 0); break;
        case "End": moveActive(current.row, cols - 1); break;
        case "PageUp": moveActive(current.row - pageSize, current.col); break;
        case "PageDown": moveActive(current.row + pageSize, current.col); break;
        default: return; // 其它按键不拦截
      }
      event.preventDefault();
    },
    [active, cols, empty, getCell, moveActive, stickyH, viewport.h],
  );

  // 行组件共用属性：引用稳定，memo 才能挡住滚动带来的整体重渲染
  const sharedRowProps = useMemo(
    () => ({ colOffsets, colWidths, getCell, onCellActivate: handleCellActivate }),
    [colOffsets, colWidths, getCell, handleCellActivate],
  );

  /* -------------------------------- 渲染 -------------------------------- */

  if (empty) {
    return <div className="flex h-full w-full items-center justify-center bg-app text-[12px] text-faint">空工作表</div>;
  }

  return (
    <div className="relative h-full w-full overflow-hidden bg-app">
      <div
        ref={scrollerRef} tabIndex={0} role="grid" aria-label="表格" aria-rowcount={rows} aria-colcount={cols}
        onScroll={handleScroll} onKeyDown={handleKeyDown}
        className="relative h-full w-full overflow-auto bg-app outline-none"
      >
        {/* 列字母表头：整块 sticky top 冻结，内部左上角单元格再 sticky left 冻结 */}
        <div className="sticky top-0 z-20 border-b border-line bg-panel" style={{ width: contentWidth, height: HEADER_H }}>
          {visibleCols.map((col) => (
            <div
              key={col} role="columnheader" aria-colindex={col + 1}
              className={`absolute top-0 truncate border-r border-line text-center text-[11px] ${
                active && active.col === col ? "bg-accent-soft text-accent" : "bg-panel text-muted"
              }`}
              style={{ left: colOffsets[col], width: colWidths[col], height: HEADER_H, lineHeight: `${HEADER_H}px` }}
            >
              {columnLabel(col)}
            </div>
          ))}
          <div className="sticky left-0 z-30 border-r border-line bg-panel" style={{ width: GUTTER_W, height: HEADER_H }} />
        </div>

        {/* 冻结的数据表头行：占一行流式高度，并 sticky 在列字母表头下方 */}
        {headerRow && rows > 0 ? (
          <SheetRow
            {...sharedRowProps}
            row={0} label="1" top={HEADER_H} width={contentWidth} frozen
            firstCol={viewRange.firstCol} lastCol={viewRange.lastCol}
            activeCol={active && active.row === 0 ? active.col : -1}
          />
        ) : null}

        {/* 滚动区：高度 = 行数 × 行高；行元素按 top 绝对定位。
            背景画一层淡行线：快速拖动时尚未渲染的区域显示为「空表格」而不是一片白 */}
        <div
          className="relative"
          style={{
            width: contentWidth,
            height: scrollRows * ROW_H,
            contain: "layout paint",
            backgroundImage: `repeating-linear-gradient(to bottom, var(--color-line) 0 1px, transparent 1px ${ROW_H}px)`,
          }}
        >
          {visibleRows.map((row) => (
            <SheetRow
              {...sharedRowProps}
              key={row}
              row={row} label={String(row + 1)} top={(row - headerOffset) * ROW_H} width={contentWidth} frozen={false}
              firstCol={viewRange.firstCol} lastCol={viewRange.lastCol}
              activeCol={active && active.row === row ? active.col : -1}
            />
          ))}
        </div>
      </div>

      {/* 徽标浮层：放在滚动容器之外，不随内容滚动 */}
      {readOnly || truncated ? (
        <div className="pointer-events-none absolute right-2 top-0 z-40 flex items-center gap-1" style={{ height: HEADER_H }}>
          {truncated && (
            <span className="rounded border border-line bg-elevated px-1.5 py-0.5 text-[10.5px] text-warning">已截断，仅显示部分数据</span>
          )}
          {readOnly && (
            <span className="rounded border border-line bg-elevated px-1.5 py-0.5 text-[10.5px] text-muted">只读</span>
          )}
        </div>
      ) : null}
    </div>
  );
}
