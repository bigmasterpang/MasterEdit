/**
 * 虚拟化表格网格（Excel 查看器风格），0.19.0 起在只读查看之上增加**可选**的单元格内联编辑。
 * - 行虚拟化：外层滚动容器 + 内层占位块（高度 = 各行高度之和；未开启换行时 = 行数 × ROW_H），
 *   只渲染与视口相交的行（上下各加 OVERSCAN 行缓冲），行绝对定位（top 由行高索引给出），
 *   百万行也无压力；开启自动换行后行高可变，虚拟化与滚动定位一律走同一个索引。
 * - 冻结装饰：列字母表头 sticky top、行号栏 sticky left、左上角单元格两个方向都 sticky。
 * - 列宽：采样前 60 行估算，memo 只依赖「列数 / 总行数 / getCell 引用」，滚动不重采样。
 * - 编辑（仅在 editable=true 时才会多出任何 DOM / 事件 / 分支；否则与改造前逐字节一致）：
 *   · 进入：双击（全选）／F2／Enter（光标在末尾）／直接敲可打印字符（该字符替换原值）。
 *     编辑框是覆盖在该格上、绝对定位在**行容器**内的非受控 <input>：跟随行一起虚拟化与滚动，
 *     不需要额外的坐标换算，退出编辑时也从 DOM 直接读值。
 *   · 提交：Enter↓ / Shift+Enter↑ / Tab→ / Shift+Tab← / 失焦 / 滚动（后两者原地不动），
 *     一律回调 onCellCommit(row, col, text)，再按方向 moveActive + ensureVisible。
 *   · 取消：Esc 丢弃本次输入（不回调 onCellCommit），保持选中。
 *   · 编辑中的按键全部由 input 自己处理，网格键盘导航在 editingRef 上提前返回。
 * - 选区与剪贴板（0.20.0，只读下同样可用：选区、复制、清空的**写回**由父组件决定）：
 *   · 状态机：`active`（活动格 = 键盘落点）+ `anchor`（选区固定角）+ `range`（渲染用矩形）。
 *     点选/键盘移动 → 选区收缩成活动格本身的 1×1（**不回调 null**，只有父组件显式传
 *     selection={null} 或表格为空时才为 null）；Shift+点击 / Shift+方向键 / Shift+Home/End /
 *     Shift+PageUp/Down → 以锚点拉到现在这一格；整行/整列/全表同理。
 *   · 拖拽框选：单元格按下后在 window 上跟踪 mousemove（React 行内只负责按下），
 *     指针位置换算成行列（列用二分查找），鼠标抬起触发一次 onSelectionCommit。
 *     拖拽期间活动格固定在锚点、**不逐帧回调 onActiveCell**，只回调 onSelectionChange。
 *     指针贴到视口边缘时按帧自动滚动（仅在拿到真实布局时生效）。
 *   · 渲染：选区内每行画一层绝对定位覆盖层（浅底 + 外边 2px 强调色边框，box-sizing 为
 *     border-box 所以不改变任何布局/行高）。行组件只拿到 4 个原始值 props：
 *     selStartCol / selEndCol（-1 = 本行不在选区）/ selTop / selBottom —— 原始值比较天然
 *     稳定，拖拽时只有「选区端点真正变化的那几行」重渲染（不引入对象切片缓存）。
 *   · 键盘：Delete/Backspace → onClearSelection；Ctrl+C → onCopySelection；Ctrl+V →
 *     navigator.clipboard.readText() 后 onPaste；Ctrl+X = 复制 + 清空；Ctrl+A → 全选。
 *     父组件没接 onCopySelection 时退回网格内部把选区序列化成 TSV 写剪贴板（兼容旧调用方）。
 * - 性能隔离（详见 pendingRowSlices / sharedRowProps 处注释）：
 *   1) 编辑期间击键只在**编辑框组件内部**更新状态（受控文本，公式补全需要当前文本与光标），
 *      网格行完全不参与，进入/退出编辑各渲染一次，memo 过的 SheetRow 只让「涉及编辑的那一行」重渲染。
 *   2) pendingCells 不整表下传，而是切成「列→文本」的行级切片：内容没变的行复用上一次的
 *      切片对象，父组件每次换新 Map 时其它行的 props 引用不变，memo 直接跳过。
 *   3) 选区同理：只传行级原始值（见上），指针在同一格内移动不会产生任何更新。
 * - 公式补全（0.20.0）：编辑框里输入 `=` 开头的内容时，由父组件注入的 completeFormula 给出候选，
 *   面板 fixed 定位在编辑框下方 —— 因为网格内容区带 `contain: layout paint`（fixed 的包含块会变成它，
 *   且会被裁剪），所以面板用 createPortal 挂到 document.body 上，保证不被滚动容器裁剪。
 *   面板打开时接管 ↑↓（循环移动高亮）、Tab/Enter（补全，光标落在括号后）、Esc（只关面板不退出编辑）；
 *   面板关闭时这些键的行为与原来完全一致。
 * - 自动列宽 / 自动换行 / 自动行高（0.21.0，只在传了 autoFitToken / wrapText 时才生效）：
 *   · autoFitToken 变化 → 按**已加载**内容采样一次，算出每列合适宽度并覆盖采样列宽（自动调整列宽）；
 *   · wrapText 开启 → 单元格文本按列宽折行（white-space: pre-wrap + word-break: break-word），
 *     行高 = 折行数 × ROW_H，由行高索引统一驱动虚拟化、定位、选区外框与编辑框；
 *   · 关闭时一切与改动前逐字节一致（仍然 truncate、固定 ROW_H），只读路径与老断言不受影响。
 *   可变行高的实现见 buildRowIndex（默认行高 + 例外行的稀疏前缀和，百万行也不会退化）。
 *   · 手动调整列宽 / 行高（0.21.0）：拖动列标头右边界 / 行号栏下边界改尺寸，双击边界按内容自动调整；
 *     受控（父组件传了 columnWidths / rowHeights）时网格只回调 onColumnResize / onRowResize，
 *     值由父组件持有并回传（这样调整列宽/行高也能进撤销栈）；未受控时留在网格会话内。
 *   · 未渲染区域的占位行线只在「全表默认行高」时画：固定 24px 纹理一旦遇到非默认行高
 *     （折行撑高、手动/受控行高）就与真实行边界错位，此时宁可不画，行边界交给每行自己的 border-b。
 *   · 一键自动调整行高（autoFitRowsToken，与 autoFitToken 对称）：把扫描窗口内**已加载**的行按内容
 *     重算行高（未加载行保持默认、关闭换行时恒为默认），所有变化收集成**一批**回调 onRowsResize，
 *     父组件一次 pushLayoutStep 就能记成一个撤销步；没传该回调时退回逐行 onRowResize。
 */
import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { JSX } from "react";
import { createPortal } from "react-dom";
import type { SheetCell, SheetCellType } from "../../types";
import { applyFormulaSuggestion, type FormulaFunction, type FormulaSuggestion } from "../../utils/formulaFunctions";
import { FormulaSuggestPanel } from "./FormulaSuggest";

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
/** 拖拽框选贴到视口左右边缘时的自动滚动步长（垂直方向按当前第一可见行的高度走） */
const AUTO_SCROLL_STEP_X = 48;
/** 12px 字号下单字符的估算宽度（列宽采样、自动列宽与折行估算共用） */
const CHAR_W = 7.2;
/** 单元格左右内边距 + 右边框（px-1.5 = 6px × 2，再加 1px 右边线）：内容可用宽度 = 列宽 - 它 */
const CELL_PAD_X = 13;
/** 自动调整列宽的夹取上限（比默认采样的 320 宽：它是用户显式要求的「按内容撑开」） */
const MAX_AUTOFIT_COL_W = 480;
/** 自动调整列宽最多采样多少行（百万行的工作簿不可能全表扫描） */
const AUTOFIT_SAMPLE_ROWS = 200;
/** 单行折行的行数上限：极长文本不至于把一行撑到几十屏，也是行高索引的保护 */
const WRAP_MAX_LINES = 80;
/** 折行高度最多扫描多少列（默认区域 50 列；超宽工作表按最左 256 列估算） */
const WRAP_SCAN_COLS = 256;
/** 行高扫描在渲染窗口之外再多算几行，减少「滚动时行高才被修正」的抖动 */
const HEIGHT_SCAN_MARGIN = 8;
/** 公式引用框的颜色（多个引用循环取色，接近 Excel 里不同引用不同颜色的做法） */
const FORMULA_TONES = ["border-accent", "border-warning", "border-danger"];
/**
 * 手动调整列宽 / 行高的夹取范围与命中区厚度。
 * **手动值优先于采样值与自动调整结果，且只在会话内有效**（暂不写回 xlsx，持久化是后续单独一件事）。
 */
const MIN_MANUAL_COL_W = MIN_COL_W;
const MAX_MANUAL_COL_W = 1200;
const MIN_MANUAL_ROW_H = ROW_H;
const MAX_MANUAL_ROW_H = 1200;
/** 列/行边界的命中区厚度（px）：贴边拖动才不费劲，又不至于挡住正常点选 */
const RESIZE_HIT = 6;
/** 拖拽填充柄（Excel 的小方块）的边长（px） */
const FILL_HANDLE = 7;

/**
 * 拖拽填充的目标区域：从 source 出发，按指针位置**单向**扩展（Excel 语义）。
 * 只扩行或只扩列，取超出的格数更多的那一维；指针落在 source 内部时原样返回（没有扩展）。
 */
export function extendFillTarget(source: SheetRange, point: { row: number; col: number }): SheetRange {
  const down = Math.max(0, point.row - source.endRow);
  const up = Math.max(0, source.startRow - point.row);
  const right = Math.max(0, point.col - source.endCol);
  const left = Math.max(0, source.startCol - point.col);
  const vertical = Math.max(down, up);
  const horizontal = Math.max(right, left);
  if (vertical === 0 && horizontal === 0) return source;
  if (vertical >= horizontal) {
    return down >= up
      ? { ...source, endRow: source.endRow + down }
      : { ...source, startRow: source.startRow - up };
  }
  return right >= left
    ? { ...source, endCol: source.endCol + right }
    : { ...source, startCol: source.startCol - left };
}

type GetCell = (row: number, col: number) => SheetCell | undefined;

/** 矩形选区（含端点，数据行号/列号） */
export interface SheetRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/** 右键目标：单元格 / 行号栏 / 列标头（row/col 为 0 起数据坐标） */
export interface SheetContextTarget {
  row: number;
  col: number;
  kind: "cell" | "rowHeader" | "colHeader";
}

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
  /** 是否允许编辑（默认 false = 纯只读，行为与现在完全一致） */
  editable?: boolean;
  /**
   * 待提交的编辑：key = `${row},${col}` → 该格当前的编辑文本（覆盖原始显示值）。
   * 父组件持有真正的编辑数据，网格只负责显示与交互。
   * 只读（editable=false）时这张表完全不参与渲染，保证只读路径零回归。
   */
  pendingCells?: ReadonlyMap<string, string>;
  /** 单元格编辑提交（Enter/Tab/失焦）。父组件负责把编辑写入自己的待提交列表 */
  onCellCommit?: (row: number, col: number, text: string) => void;
  /** 受控选区；null/undefined 表示只有活动单元格没有选区。不传 = 非受控（网格自己管） */
  selection?: SheetRange | null;
  /** 选区变化（拖拽中会高频触发，父组件自己做节流/汇总） */
  onSelectionChange?: (range: SheetRange | null) => void;
  /** 选区变化结束（鼠标抬起 / 键盘调整后），父组件用它算统计、避免拖拽中反复计算 */
  onSelectionCommit?: (range: SheetRange | null) => void;
  /** Delete / Backspace：清空选区（父组件决定怎么把空格写回） */
  onClearSelection?: (range: SheetRange) => void;
  /** Ctrl+C：复制选区（父组件负责序列化为 TSV 并写剪贴板） */
  onCopySelection?: (range: SheetRange) => void;
  /** Ctrl+V：粘贴（网格负责读剪贴板文本后回调，父组件解析并批量提交） */
  onPaste?: (range: SheetRange, text: string) => void;
  /**
   * 单元格/表头右键：父组件负责弹菜单；row/col 为 0 起数据坐标。
   * 目标不在当前选区内时，网格会先按 Excel 规则把选区/活动格移到该目标，再回调；
   * 传了本回调才会 preventDefault 掉系统菜单（没传则保留浏览器默认菜单）。
   */
  onCellContextMenu?: (target: SheetContextTarget, position: { x: number; y: number }) => void;
  /**
   * 公式补全：输入 `=` 后由父组件注入的匹配函数；不传则不提示。
   * 网格只负责「显示候选 + 键盘/鼠标选中 + 插入文本」，匹配规则完全由父组件决定。
   */
  completeFormula?: (text: string, caret: number) => { from: number; to: number; items: FormulaFunction[] } | null;
  /**
   * 自增令牌：值变化时做一次「按内容自动调整列宽」（Excel 的自动调整列宽）。
   * 只基于**已加载**的内容采样（见 computeAutoFitWidths 的取舍说明）；不传这个 prop 就完全不做自动列宽，
   * 列宽仍走原来的采样逻辑。
   */
  autoFitToken?: number;
  /**
   * 自动换行：单元格文本超过列宽时折行显示，行高按折行数自动变高（可变行高）。
   * 只读网格同样生效（只读也要能看全内容）；关闭时与改动前的单行 + truncate 行为逐字节一致。
   */
  wrapText?: boolean;
  /** 列宽/行高变化导致内容总尺寸变化后回调（父组件据此刷新滚动条、统计等；不传也没关系） */
  onLayoutChange?: () => void;
  /**
   * 外部（内容栏 / 公式栏）正在编辑公式：为 true 时在网格上按下/拖动**不夺焦**、不进入编辑，
   * 松开鼠标后把选中的区域通过 onPickReference 回调出去（引用文本由父组件自己转换）。
   * 与「单元格内编辑框里输入 = 后拖选引用」共用同一套拖选代码，区别只是谁消费结果。
   */
  pickReference?: boolean;
  /** 外部拾取结束：把选中的区域回调出去（数据坐标，0 起、含端点；单格也会回调） */
  onPickReference?: (range: SheetRange) => void;
  /**
   * 当前单元格公式引用的区域（0 起、含端点）：网格用彩色虚线框标出「公式的作用范围」。
   * 父组件负责解析引用（`B:B` / `2:2` 展开、跨表判断等）；网格只负责画框：
   * 只画落在可视窗口内的部分（不会因此多渲染任何行），不拦截鼠标，也不改变选区样式。
   */
  formulaRanges?: ReadonlyArray<SheetRange>;
  /**
   * 手动列宽覆盖（列索引 → 像素），由父组件持有 —— 传了（哪怕空对象）就进入**受控**模式：
   * 网格不再自己记一份手动列宽，拖动/双击只回调 onColumnResize，宽度以这里传回来的值为准。
   * 这样父组件才能把「调整列宽」记进撤销栈（撤销 = 把这张表改回旧值）。
   * 优先级最高：覆盖采样宽度与 autoFitToken 的自动调整结果（自动调整只影响没被覆盖的列）。
   */
  columnWidths?: Readonly<Record<number, number>>;
  /** 手动行高覆盖（行索引 → 像素），语义同 columnWidths：受控、优先、可撤销恢复 */
  rowHeights?: Readonly<Record<number, number>>;
  /** 用户手动调整列宽（拖动中每帧 + 双击边界各回调一次）：父组件更新 columnWidths 让它生效 */
  onColumnResize?: (index: number, width: number) => void;
  /** 用户手动调整行高（拖动中每帧 + 双击边界各回调一次）：父组件更新 rowHeights 让它生效 */
  onRowResize?: (index: number, height: number) => void;
  /**
   * 自增令牌：值变化时把所有**已加载**的行按内容自动调整行高（Excel 的「自动调整行高」，
   * 与 autoFitToken 列宽那一个对称）。只处理「渲染窗口 + overscan + 余量」里当前已加载的行：
   * 不扫全表，也不会把这一次动作记在尚未加载的行上（未加载行保持默认，等数据到达后由换行扫描按需修正）。
   * 关闭 wrapText 时行高恒为默认（单行），不会因为一键调整把行撑高；同一令牌值只处理一次。
   */
  autoFitRowsToken?: number;
  /**
   * 一键自动调整行高产生的**一批**行高变化：一次调用给出本次所有变化的行（长度 = 真正变化的行数），
   * 父组件用一次 pushLayoutStep 就能把它记成一个撤销步。没传这个回调时退回逐行 onRowResize。
   */
  onRowsResize?: (changes: Array<{ index: number; height: number }>) => void;
  /**
   * 一键自动调整列宽产生的**一批**列宽变化（可选，形状与 onRowsResize 对称）：
   * 传了就用它（父组件一次记成一个撤销步），没传则逐列退回 onColumnResize。
   * 只在**用户显式触发**（autoFitToken 变化）时回调；挂载时那次自动拟合不回调，避免凭空多出撤销步。
   */
  onColumnsResize?: (changes: Array<{ index: number; width: number }>) => void;
  /**
   * 冻结的行数（0 = 不冻结；1 = 冻结首行）：冻结行在垂直滚动时始终可见，其余行正常滚动。
   * 与 headerRow 是同一套实现（headerRow 相当于「冻结 1 行 + 表头样式」），同时传时取两者较大的行数，
   * 不会出现两层表头；冻结区与滚动区之间会多画一条分隔线。
   */
  freezeRows?: number;
  /** 冻结的列数（0 = 不冻结；1 = 冻结首列）：冻结列在水平滚动时始终可见（行号栏永远是最左边那一列） */
  freezeCols?: number;
  /**
   * 查找命中的单元格（0 起）：在格子上铺一层**比选区更浅**的提示底色，用于「查找命中」。
   * 只影响渲染窗口内的格子（不额外渲染行），不改选区、不改活动单元格；与选区重叠时**选区优先**。
   * 不传 / 空数组时 DOM 与改动前逐字节一致。
   */
  findHits?: ReadonlyArray<{ row: number; col: number }>;
  /** 当前定位到的那条命中（面板上的「当前/总数」）：底色略强并加一圈细描边，与其它命中区分 */
  currentHit?: { row: number; col: number } | null;
  /**
   * 请求把某格滚动到可见：**token 变化**时执行一次（同一个 token 只处理一次）。
   * 只滚动，不改选区 / 活动单元格 / 焦点框 —— 查找跳转时用户的选区高亮必须保持不动。
   * 目标落在冻结区里时不滚动（与 ensureVisible 的既有语义一致）；不传时行为与改动前一致。
   */
  scrollTarget?: { row: number; col: number; token: number } | null;
  /**
   * 拖拽自动填充（Excel 的填充柄）完成时回调：
   * `source` 是原选区（提供规律样本），`target` 是填充后的完整区域（**一定包含 source**，
   * 且只朝行方向或列方向单向扩展）。规律推断与写入由父组件负责（见 utils/fillSeries.ts）。
   * **不传这个回调就完全不渲染填充柄**（DOM 与改动前逐字节一致）。
   */
  onFillRange?: (source: SheetRange, target: SheetRange) => void;
}

/** 公式补全匹配器（从 props 里派生，保证与 props 签名永远一致） */
type FormulaCompleter = NonNullable<SheetGridProps["completeFormula"]>;

/**
 * 调父组件注入的匹配器，并把异常吞掉：补全只是锦上添花，
 * 绝不能让匹配器里的一个 TypeError 把整张网格（乃至整个视图）带崩。
 */
function safeComplete(
  completer: FormulaCompleter,
  text: string,
  caret: number,
): ReturnType<FormulaCompleter> {
  try {
    return completer(text, caret);
  } catch {
    return null;
  }
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
    widths[col] = Math.min(MAX_COL_W, Math.max(MIN_COL_W, 18 + Math.round(longest[col] * CHAR_W)));
  }
  return widths;
}

/* ------------------------------ 公式引用（拾取） ------------------------------ */

/** 单元格引用：可选的列绝对符 + 列名 + 可选的行绝对符 + 行号（$B$3 / B3 / B$3 / $B3） */
const REF_CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
/** 整列 / 整行引用（B:B / $B:$D / 2:2） */
const REF_COLS = String.raw`\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}`;
const REF_ROWS = String.raw`\$?\d+:\$?\d+`;

/**
 * 光标前「紧邻的引用记号」：用于判断这次拾取是**替换**还是**插入**。
 * 覆盖 B3 / $B$3 / B3:B4 / $B$3:$C$4 / B:B / 2:2 这些写法（与 Excel 一致）。
 */
const REFERENCE_TAIL = new RegExp(`(?:${REF_CELL}:${REF_CELL}|${REF_CELL}|${REF_COLS}|${REF_ROWS})$`);

/**
 * 选区 → 公式里的引用文本（Excel 写法）：单格 `B3`、矩形 `B3:B4`、整列 `B:B`、整行 `3:3`。
 * 行号按 Excel 习惯从 1 开始（数据第 0 行 = 第 1 行）。
 */
function rangeReference(range: SheetRange, mode: "cell" | "col" | "row"): string {
  if (mode === "col") return `${columnLabel(range.startCol)}:${columnLabel(range.endCol)}`;
  if (mode === "row") return `${range.startRow + 1}:${range.endRow + 1}`;
  const start = `${columnLabel(range.startCol)}${range.startRow + 1}`;
  const end = `${columnLabel(range.endCol)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/**
 * 拖动指示线的落点：client 坐标 → **网格容器内部坐标**。
 * 指示线是绝对定位在网格容器（.relative）里的，而容器不一定在视口原点
 * （网格上方还有标签栏 / 工具栏 / 公式栏）：直接拿 clientY 当 top 会让横线整体下移一个容器偏移量，
 * 看起来不跟着指针走。列方向同理（容器左边也可能有偏移）。
 */
function guideOffset(
  el: HTMLElement | null,
  kind: "col" | "row",
  clientX: number,
  clientY: number,
): number {
  const rect = el?.getBoundingClientRect();
  return kind === "col" ? clientX - (rect?.left ?? 0) : clientY - (rect?.top ?? 0);
}

/** 手动列宽的夹取（父组件传进来的受控值同样夹一遍：非法值不至于把布局撑爆） */
function clampManualColWidth(width: number | undefined): number | null {
  if (width === undefined || !Number.isFinite(width)) return null;
  return Math.max(MIN_MANUAL_COL_W, Math.min(MAX_MANUAL_COL_W, Math.round(width)));
}

/** 手动行高的夹取（同上） */
function clampManualRowHeight(height: number | undefined): number | null {
  if (height === undefined || !Number.isFinite(height)) return null;
  return Math.max(MIN_MANUAL_ROW_H, Math.min(MAX_MANUAL_ROW_H, Math.round(height)));
}

/**
 * 受控覆盖表（columnWidths / rowHeights）→ Map<索引, 像素>：只认非负整数键，非法项直接忽略。
 * 每次 props 换新对象都会重建（撤销就是把旧对象传回来），因此调用方把它放进 useMemo 依赖即可。
 */
function controlledSizes(record: Readonly<Record<number, number>> | undefined): Map<number, number> | null {
  if (record === undefined) return null;
  const map = new Map<number, number>();
  for (const key of Object.keys(record)) {
    const index = Number(key);
    if (!Number.isInteger(index) || index < 0) continue;
    map.set(index, record[index as unknown as number]);
  }
  return map;
}

/* ------------------------------ 自动列宽 / 换行 ------------------------------ */

/**
 * 取某一格「当前显示的内容」：**待提交的编辑优先**（用户刚输入还没保存的长字符串也要参与
 * 自动列宽与折行行高的计算，否则「输入一长串后自动列宽不生效」）。
 */
function readCellText(
  live: { getCell: GetCell; pendingCells?: ReadonlyMap<string, string> },
  row: number,
  col: number,
): string {
  return live.pendingCells?.get(`${row},${col}`) ?? live.getCell(row, col)?.v ?? "";
}

/**
 * 自动调整列宽的宽度计算（对应 Excel 的「自动调整列宽」）。
 *
 * **取舍：只基于「已加载」的内容采样。** 工作簿动辄几十万行，前端又是稀疏窗口缓存，
 * 全表扫描既没有数据也不现实，所以这里最多看 AUTOFIT_SAMPLE_ROWS 行（默认前 200 行），
 * 再补上「当前可视窗口」的那几十行 —— 用户滚到哪里，那里的内容也会参与这一次估算。
 * 未加载的行拿不到内容，自然不参与；这与原有的 computeColWidths 是同一套取舍，只是采样更多、
 * 上限更宽（用户显式点了「自动调整列宽」，就该比默认采样更舍得撑开）。
 * 内容用 readCellText 取：**待提交的编辑优先**（刚敲进去还没保存的长串同样算数）。
 *
 * 中文/全角按 2 个字符宽计（visualLength），加上左右内边距，最后夹到 [MIN_COL_W, MAX_AUTOFIT_COL_W]。
 */
/**
 * 「当前可视行」的数据行号列表（自动调整列宽的补采样用）。
 * 前 AUTOFIT_SAMPLE_ROWS 行之外的内容只有用户滚到那里才加载得到，
 * 所以这一批行必须补进去：**用户滚到哪里，那里的内容也参与这一次估算**。
 * 双击列边界（autoFitColumn）用的是完全相同的采样，两处共用这一个函数，
 * 避免「工具栏自动调整算上了屏上的内容、双击边界却算法不同」的偏差。
 */
function visibleRowList(
  viewRange: { firstVisible: number; lastVisible: number },
  headerOffset: number,
): number[] {
  const rows: number[] = [];
  for (let v = viewRange.firstVisible; v <= viewRange.lastVisible; v += 1) rows.push(v + headerOffset);
  return rows;
}

/**
 * 行高处理的窗口（数据行号，含端点）：渲染窗口（已经含 overscan）+ 上下各 HEIGHT_SCAN_MARGIN 的余量。
 * 换行折行高度扫描与「一键自动调整行高」共用同一个窗口 —— 两者都不扫全表（百万行工作表不可能）。
 */
function heightScanWindow(
  viewRange: { firstRow: number; lastRow: number },
  headerOffset: number,
  rows: number,
): { first: number; last: number } {
  return {
    first: Math.max(0, viewRange.firstRow + headerOffset - HEIGHT_SCAN_MARGIN),
    last: Math.min(rows - 1, viewRange.lastRow + headerOffset + HEIGHT_SCAN_MARGIN),
  };
}

function computeAutoFitWidths(
  cols: number,
  sampleRows: number,
  windowRows: number[],
  readText: (row: number, col: number) => string,
): number[] {
  const widths = new Array<number>(cols).fill(DEFAULT_COL_W);
  if (cols === 0) return widths;
  const longest = new Array<number>(cols).fill(0);
  const scan = (row: number) => {
    for (let col = 0; col < cols; col += 1) {
      const text = readText(row, col);
      if (text.length === 0) continue;
      const length = visualLength(text);
      if (length > longest[col]) longest[col] = length;
    }
  };
  for (let row = 0; row < sampleRows; row += 1) scan(row);
  for (const row of windowRows) if (row >= sampleRows) scan(row);
  for (let col = 0; col < cols; col += 1) {
    if (longest[col] === 0) continue;
    widths[col] = Math.min(MAX_AUTOFIT_COL_W, Math.max(MIN_COL_W, 18 + Math.round(longest[col] * CHAR_W)));
  }
  return widths;
}

/**
 * 文本测量：折行行数直接决定行高，**估算偏小就会把内容裁掉**，所以宽度只能偏大不能偏小。
 * 优先用 canvas 按单元格的真实字体逐字测宽（结果按「字体 + 字符」缓存，每个字符整表只测一次）；
 * 拿不到 canvas（jsdom / SSR）时退化成保守的字符宽度表。
 * 逐字相加天然是上界：真实排版有字距调整（kerning）与连字，实际宽度只会更窄。
 */
let measureCache: { font: string; ctx: CanvasRenderingContext2D; widths: Map<string, number> } | null = null;
let measureUnavailable = false;

/** 没有 canvas 时的保守字符宽度（12px 字号的上界）：CJK/全角按字号整宽，西文按宽字形上界 */
const FALLBACK_WIDE_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@%&MWmw";
const FALLBACK_NARROW_CHARS = "ilj.,:;'|!()[]{}ftr/- ";
function fallbackCharWidth(ch: string): number {
  const code = ch.codePointAt(0) ?? 0;
  if (code > 0x2e7f) return 12; // CJK / 全角
  if (FALLBACK_WIDE_CHARS.includes(ch)) return 9.5; // 宽西文 / 数字
  if (FALLBACK_NARROW_CHARS.includes(ch)) return 6; // 窄字符
  return 8.5;
}

/** 单字符宽度（px）。1.02 是四舍五入与字距的安全系数，宁可多算一点 */
function charWidth(ch: string, font: string): number {
  if (!measureUnavailable) {
    if (!measureCache || measureCache.font !== font) {
      let ctx: CanvasRenderingContext2D | null = null;
      try {
        ctx = typeof document !== "undefined" ? document.createElement("canvas").getContext("2d") : null;
      } catch {
        ctx = null; // 某些环境（jsdom 没装 canvas 包等）会直接抛，退化成保守表
      }
      if (!ctx) {
        measureUnavailable = true;
      } else {
        ctx.font = font;
        measureCache = { font, ctx, widths: new Map() };
      }
    }
    const cache = measureCache;
    if (cache && cache.font === font) {
      const cached = cache.widths.get(ch);
      if (cached !== undefined) return cached;
      const measured = cache.ctx.measureText(ch).width;
      const value = measured > 0 ? measured * 1.02 : fallbackCharWidth(ch);
      cache.widths.set(ch, value);
      return value;
    }
  }
  return fallbackCharWidth(ch);
}

/** 一段文本的宽度（逐字相加，上界） */
function textWidth(text: string, font: string): number {
  let total = 0;
  for (const ch of text) total += charWidth(ch, font);
  return total;
}

/**
 * 折行后的行数：按 CSS `white-space: pre-wrap; overflow-wrap: break-word` 的行为模拟 ——
 * 先按词边界换行（空格处断），词本身比一行还宽时再按字符切开；显式换行符单独起段。
 * 宽度取上界 + 逐行贪心填充 + 每段至少 1 行，因此结果 ≥ 浏览器实际需要的行数。
 */
function wrapLineCount(text: string, contentWidth: number, font: string): number {
  if (text.length === 0) return 1;
  const limit = Math.max(1, contentWidth);
  let lines = 0;
  for (const segment of text.split("\n")) {
    if (segment.length === 0) {
      lines += 1;
      continue;
    }
    let count = 1;
    let used = 0;
    for (const token of segment.match(/\S+\s*/g) ?? [segment]) {
      const width = textWidth(token, font);
      if (used + width <= limit) {
        used += width;
        continue;
      }
      if (used > 0) {
        count += 1;
        used = 0;
      }
      if (width <= limit) {
        used = width;
        continue;
      }
      // 词本身放不下一行：break-word 会按字符切开
      for (const ch of token) {
        const w = charWidth(ch, font);
        if (used + w > limit && used > 0) {
          count += 1;
          used = 0;
        }
        used += w;
      }
    }
    lines += count;
    if (lines >= WRAP_MAX_LINES) return WRAP_MAX_LINES;
  }
  return Math.max(1, lines);
}

/**
 * 行高索引：把「固定行高」换成「默认行高 + 例外行」的稀疏前缀和，用于 scrollTop ↔ 行号互转。
 *
 * 为什么不用稠密前缀和数组：行数可达 1,048,576（xlsx 上限），而内容只有滚动到的那几十行是**已加载**的。
 * 稠密数组既要给未加载的行编造高度（前端根本不知道），又要在每次数据/列宽变化时重算上百万项。
 * 稀疏表只记录「高度 ≠ ROW_H 的已加载行」，按虚拟行号升序存下 [行号, 高度, 该行顶部 y]：
 *   - topOf(row)：二分找到它前面最近的例外行，再按默认行高线性外推 —— O(log n)，n = 例外行数；
 *   - virtualAt(y)：二分找到 y 落在哪个例外行之后，再算 (y - yStart) / ROW_H，并夹到下一个例外行之前
 *     （例外行内部高度不同，不能继续按默认行高外推）；
 *   - totalHeight = 所有行高度之和（滚动条总高度）。
 * 行数上限 1000×50 的默认区域下，例外行数最多几百条，重建是一次几百项的排序，不存在 O(n²)。
 * 例外行数 = 「滚动看过的窗口里折了行的行数」之和（几千条量级）；重建只在行高/列宽/数据版本变化时发生，
 * 滚动本身不会重建索引，也不会做 O(行数) 的工作。
 */
interface RowIndex {
  /** 某一数据行的行高（未加载 / 未换行时就是 ROW_H） */
  heightOf(row: number): number;
  /** 某一数据行在滚动内容里的顶部 y（不含顶部冻结区） */
  topOf(row: number): number;
  /** 内容坐标 y（不含冻结区）→ 虚拟行号（0 = 第一个可滚动行） */
  virtualAt(y: number): number;
  /** 全部可滚动行的高度之和 = 滚动条内容总高度 */
  totalHeight: number;
}

function buildRowIndex(rows: number, headerOffset: number, heights: ReadonlyMap<number, number>): RowIndex {
  const scrollRows = Math.max(0, rows - headerOffset);
  /** [虚拟行号, 该行高度, 该行顶部 y]，按虚拟行号升序 */
  const marks: Array<[number, number, number]> = [];
  let y = 0;
  let cursor = 0;
  for (const row of [...heights.keys()].sort((a, b) => a - b)) {
    const v = row - headerOffset;
    const h = heights.get(row) ?? ROW_H;
    if (v < 0 || v >= scrollRows || h === ROW_H) continue;
    y += (v - cursor) * ROW_H;
    marks.push([v, h, y]);
    y += h;
    cursor = v + 1;
  }
  const totalHeight = y + (scrollRows - cursor) * ROW_H;

  const heightOf = (row: number): number => {
    const h = heights.get(row);
    return h === undefined || h === ROW_H ? ROW_H : h;
  };

  /** 最后一个「虚拟行号 ≤ v」的例外行下标（没有则 -1） */
  const markBeforeRow = (v: number): number => {
    let lo = 0;
    let hi = marks.length - 1;
    let found = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (marks[mid][0] <= v) {
        found = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return found;
  };

  /** 最后一个「顶部 y ≤ value」的例外行下标（没有则 -1） */
  const markBeforeY = (value: number): number => {
    let lo = 0;
    let hi = marks.length - 1;
    let found = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (marks[mid][2] <= value) {
        found = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return found;
  };

  return {
    totalHeight,
    heightOf,
    topOf(row: number): number {
      const v = row - headerOffset;
      if (v <= 0) return 0;
      const i = markBeforeRow(v);
      if (i < 0) return v * ROW_H;
      const [mv, mh, my] = marks[i];
      if (v === mv) return my;
      // 例外行之后要跳过它自己的高度（它是 mh 而不是 ROW_H）
      return my + mh + (v - mv - 1) * ROW_H;
    },
    virtualAt(value: number): number {
      if (scrollRows <= 0) return 0;
      const clamped = value <= 0 ? 0 : value;
      const i = markBeforeY(clamped);
      let v: number;
      if (i < 0) {
        v = Math.floor(clamped / ROW_H);
      } else {
        const [mv, mh, my] = marks[i];
        // y 落在例外行内部就是它自己；否则从它的下边界开始按默认行高外推
        v = clamped < my + mh ? mv : mv + 1 + Math.floor((clamped - my - mh) / ROW_H);
      }
      const nextV = i + 1 < marks.length ? marks[i + 1][0] : scrollRows;
      if (v >= nextV) v = nextV - 1;
      return v < 0 ? 0 : v > scrollRows - 1 ? scrollRows - 1 : v;
    },
  };
}

/* -------------------------------- 选区工具 -------------------------------- */
/** 夹取到 [0, max] */
function clampIndex(value: number, max: number): number {
  return value < 0 ? 0 : value > max ? max : value;
}

/** 两个角点 → 归一化矩形（start ≤ end） */
function rangeOf(a: { row: number; col: number }, b: { row: number; col: number }): SheetRange {
  return {
    startRow: Math.min(a.row, b.row),
    startCol: Math.min(a.col, b.col),
    endRow: Math.max(a.row, b.row),
    endCol: Math.max(a.col, b.col),
  };
}

/** 单格选区 */
function singleRange(row: number, col: number): SheetRange {
  return { startRow: row, startCol: col, endRow: row, endCol: col };
}

/** 两个选区是否等价（含 null）；用于「值没变就不重渲染 / 不回调」 */
function sameRange(a: SheetRange | null, b: SheetRange | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.startRow === b.startRow && a.startCol === b.startCol && a.endRow === b.endRow && a.endCol === b.endCol;
}

/** 某格是否落在选区内（含端点） */
function inRange(range: SheetRange | null, row: number, col: number): boolean {
  return range !== null && row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol;
}

/** 把选区夹到当前数据范围内；空表或完全越界返回 null。值未变时返回原对象（引用稳定） */
function clampRange(range: SheetRange | null, rows: number, cols: number): SheetRange | null {
  if (!range || rows === 0 || cols === 0) return null;
  const next: SheetRange = {
    startRow: clampIndex(Math.min(range.startRow, range.endRow), rows - 1),
    startCol: clampIndex(Math.min(range.startCol, range.endCol), cols - 1),
    endRow: clampIndex(Math.max(range.startRow, range.endRow), rows - 1),
    endCol: clampIndex(Math.max(range.startCol, range.endCol), cols - 1),
  };
  return sameRange(range, next) ? range : next;
}

/**
 * 内容坐标 x → 列号。colOffsets 单调递增用二分（列数可能上万，不能线性扫描）；
 * x 落在行号栏上或超出最后一列时自然夹取到端点列。
 */
function colAtX(x: number, colOffsets: number[]): number {
  let lo = 0;
  let hi = colOffsets.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (colOffsets[mid] <= x) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** 选区格子总数（用于兜底复制的规模保护） */
function rangeCellCount(range: SheetRange): number {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
}

/** 拖拽框选会话：按下发生在单元格/表头上，之后的移动与抬起都监听在 window 上 */interface DragSession {
  /** 锚点（选区固定角） */
  anchor: { row: number; col: number };
  /** 拖拽模式：单元格矩形 / 整列（列标头按下）/ 整行（行号栏按下） */
  mode: "cell" | "col" | "row";
  /**
   * 公式拾取（Excel 的「点选引用」）：编辑中且文本以 = 开头（或外部内容栏正在编辑公式）时，
   * 按下鼠标不改活动格、不触发编辑，拖出来的选区在松开鼠标后被写成引用文本 / 回调给父组件。
   * `editor` = 交给单元格内编辑框，`external` = 通过 onPickReference 交给内容栏。
   */
  pick?: "editor" | "external";
  /** 最后一次指针位置（client 坐标）：自动滚动时用它重算端点 */
  clientX: number;
  clientY: number;
  /** 当前选区（同步值：mouseup 要立刻拿到，不等 React 状态回填） */
  range: SheetRange;
}

/* -------------------------------- 编辑相关 -------------------------------- */

/** 进入编辑时编辑框里的光标落点：全选（双击）或落在文本末尾（F2 / Enter / 直接输入） */
type EditCaret = "all" | "end";

/** 提交后选中格的移动方向；none = 原地不动（失焦、滚动等外部原因提交） */
type CommitMove = "down" | "up" | "right" | "left" | "none";

/** 一行待提交编辑的行级切片：列号 → 编辑文本 */
type RowPending = ReadonlyMap<number, string>;

/** 编辑态：整张网格同时只可能有一份，只有命中的那一行会拿到它的信息 */
interface EditingState {
  row: number;
  col: number;
  /** 进入编辑时的初始文本（待提交的编辑值优先，其次原始显示值） */
  text: string;
  caret: EditCaret;
}

/** 传给行组件的编辑接口：整体 memo 成稳定引用，避免行因回调重建而重渲染 */
interface RowEditorApi {
  /** editable=false 时行为组件连双击都不接（只读路径 DOM 完全不变） */
  editable: boolean;
  /** 双击进入编辑（初始内容全选） */
  startEdit: (row: number, col: number) => void;
  /** 提交编辑框里的文本并按方向移动选中格（编辑格的行列由网格自己记着） */
  commitEdit: (text: string, move: CommitMove) => void;
  cancelEdit: () => void;
  /** 把真实 input 节点交给网格（滚动提交时网格直接读它的当前值） */
  bindInput: (el: HTMLInputElement | null) => void;
  /** 读最新的公式补全匹配器（不直接传函数：父组件换函数不会改变本对象的引用，行不会白重渲染） */
  readCompleter: () => FormulaCompleter | undefined;
  /** 编辑框把「改写内容」的能力注册给网格（公式拾取要把引用写回去） */
  bindApi: (api: CellEditorApi | null) => void;
  /** 读「网格是否正在拾取」：Esc 的第一优先级是取消拾取，而不是退出编辑 */
  readPicking: () => boolean;
  /** 取消本次拾取（网格恢复拾取前的选区，编辑框不退出） */
  cancelPick: () => void;
  /** 编辑框文本是否以 = 开头（决定要不要进入拾取模式与拾取区域的视觉） */
  notifyFormulaMode: (active: boolean) => void;
}

/** 编辑框暴露给网格的能力（只有拾取模式用得到） */
interface CellEditorApi {
  /**
   * 把引用文本写进编辑框：光标前紧邻一个引用记号（B3 / $B$3 / B3:B4 / $B$3:$C$4 / B:B / 2:2）就替换它，
   * 否则在光标处插入；写完后光标停在引用之后，编辑框保持焦点（可以接着打 `)`）。
   */
  applyReference(ref: string): void;
}
interface SheetCellEditorProps {
  row: number;
  col: number;
  /** 该格的绝对左偏移与宽度（与单元格完全一致） */
  left: number;
  width: number;
  /** 行高（可变行高下编辑框要跟着行一起长高；输入框始终单行，文本在行内垂直居中） */
  height: number;
  /** 进入编辑那一刻的文本 */
  text: string;
  caret: EditCaret;
  bindInput: (el: HTMLInputElement | null) => void;
  /** 提交编辑框里的文本；move 决定提交后选中格往哪走 */
  onCommit: (text: string, move: CommitMove) => void;
  onCancel: () => void;
  /** 读最新的公式补全匹配器（走 ref，父组件每次传新函数也不会让所有行重渲染） */
  readCompleter: () => FormulaCompleter | undefined;
  /** 把「改写内容」的能力注册给网格（拾取模式要把引用写进编辑框） */
  bindApi: (api: CellEditorApi | null) => void;
  /** 读「网格是否正在拾取」（Esc 优先取消拾取） */
  readPicking: () => boolean;
  /** 取消本次拾取 */
  onPickCancel: () => void;
  /** 文本是否以 = 开头变化时通知网格 */
  onFormulaModeChange: (active: boolean) => void;
  /**
   * 正在编辑的这格是不是查找命中：编辑框是不透明的，底色不换就完全看不出命中。
   * `current` = 当前命中（面板上的「当前/总数」），用更重的底色 + 2px 高对比描边。
   */
  hitTint?: "none" | "hit" | "current";
}

/** 候选面板状态：字段与 FormulaSuggestPanel 期望的 state 一致（外加一个等价性 key） */
interface CellSuggestState {
  suggestion: FormulaSuggestion;
  index: number;
  rect: { left: number; top: number; width: number };
  /** 「替换范围 + 候选顺序」签名：没变就不重渲染 */
  key: string;
}

/**
 * 覆盖在单元格上的编辑框。
 * - 受控文本：公式补全要随时知道当前文本与光标；击键只在**编辑框组件内部**更新状态，
 *   网格行不参与渲染（配合 SheetRow 的 memo，整张网格不会因输入而重渲染）。
 * - 绝对定位在**行容器**内，因此跟随行的虚拟化与滚动，不需要额外的坐标换算。
 * - 由 SheetRow 用 `${row}:${col}` 作 key 挂载：换格即重新挂载，初值与光标落点因此总是对的。
 * - 公式候选面板用 portal 挂到 body：网格内容区带 `contain: layout paint`，
 *   它既会改变 fixed 定位的包含块、又会裁剪后代，挂在行里必然错位/被裁。
 */
function SheetCellEditor({
  row,
  col,
  left,
  width,
  height,
  text,
  caret,
  bindInput,
  onCommit,
  onCancel,
  readCompleter,
  bindApi,
  readPicking,
  onPickCancel,
  onFormulaModeChange,
  hitTint = "none",
}: SheetCellEditorProps): JSX.Element {
  const inputRef = useRef<HTMLInputElement>(null);
  /** Enter / Esc / blur 可能在一个提交里接连发生，用它保证一次编辑只提交或取消失败一次 */
  const doneRef = useRef(false);
  /** 编辑框当前文本（受控）：补全插入后要立刻反映到输入框与光标 */
  const [value, setValue] = useState(text);
  /** 公式候选面板状态（null = 不显示） */
  const [suggest, setSuggest] = useState<CellSuggestState | null>(null);
  /** 按 Esc 主动关掉面板后，本次编辑内不再自动弹出（与内容栏一致） */
  const dismissedRef = useRef(false);

  /** 按当前文本与光标重算候选；没注入匹配器 / 已关掉 / 没命中时一律不显示 */
  const refreshSuggest = useCallback(() => {
    const input = inputRef.current;
    const completer = readCompleter();
    if (!input || !completer || doneRef.current || dismissedRef.current) {
      setSuggest((prev) => (prev === null ? prev : null));
      return;
    }
    const found = safeComplete(completer, input.value, input.selectionStart ?? input.value.length);
    if (!found || found.items.length === 0) {
      setSuggest((prev) => (prev === null ? prev : null));
      return;
    }
    // 父组件只承诺 from/to/items（可能没有 query），所以用「范围 + 候选名」做等价性签名
    const key = `${found.from}:${found.to}:${found.items.map((fn) => fn.name).join(",")}`;
    setSuggest((prev) => {
      if (prev && prev.key === key) return prev; // 完全等价：不重渲染
      const rect = input.getBoundingClientRect();
      return {
        suggestion: { from: found.from, to: found.to, query: "", items: found.items },
        index: 0,
        key,
        rect: { left: rect.left, top: rect.bottom + 2, width: rect.width },
      };
    });
  }, [readCompleter]);

  // 文本变化（含补全插入）后重算候选
  useEffect(() => {
    refreshSuggest();
  }, [refreshSuggest, value]);

  /** 文本是否以 = 开头：公式拾取模式的总开关（去空白后判断，与内容栏一致） */
  const formulaMode = /^\s*=/.test(value);
  // 通知网格：进入/退出拾取模式（网格只在真的翻转时 setState）
  useEffect(() => {
    onFormulaModeChange(formulaMode);
  }, [formulaMode, onFormulaModeChange]);

  /**
   * 把引用写进编辑框（公式拾取松开鼠标时由网格调用）：
   * 光标前紧邻一个引用记号就替换它（Excel 语义：拖一次改一次范围），否则在光标处插入。
   */
  const applyReference = useCallback((ref: string) => {
    const input = inputRef.current;
    if (!input) return;
    const caret = input.selectionStart ?? input.value.length;
    const before = input.value.slice(0, caret);
    const matched = REFERENCE_TAIL.exec(before);
    const start = matched ? caret - matched[0].length : caret;
    const next = input.value.slice(0, start) + ref + input.value.slice(caret);
    const nextCaret = start + ref.length;
    setValue(next);
    // 受控值渲染完再落光标（编辑框始终保持焦点，所以接着就能打 `)`)
    window.requestAnimationFrame(() => {
      const target = inputRef.current;
      if (!target) return;
      target.focus({ preventScroll: true });
      target.setSelectionRange(nextCaret, nextCaret);
    });
  }, []);

  // 把改写能力注册给网格（拾取模式要用）；卸载时注销
  useEffect(() => {
    bindApi({ applyReference });
    return () => bindApi(null);
  }, [applyReference, bindApi]);

  // 面板打开期间跟随输入框位置（滚动会先提交编辑，这里主要应对窗口尺寸变化）
  useEffect(() => {
    if (!suggest) return;
    const onMove = () => refreshSuggest();
    window.addEventListener("resize", onMove);
    window.addEventListener("scroll", onMove, true);
    return () => {
      window.removeEventListener("resize", onMove);
      window.removeEventListener("scroll", onMove, true);
    };
  }, [refreshSuggest, suggest]);

  /** 用候选项补全：插入 `NAME(`（无参函数 `NAME()`），光标落在括号之后 */
  const pickSuggestion = (fn: FormulaFunction) => {
    const input = inputRef.current;
    if (!input || !suggest) return;
    const { text: nextText, caret: nextCaret } = applyFormulaSuggestion(input.value, suggest.suggestion, fn);
    setValue(nextText);
    setSuggest(null);
    // 受控值渲染完再落光标（与内容栏补全同一套做法）
    window.requestAnimationFrame(() => {
      const target = inputRef.current;
      if (!target) return;
      target.focus({ preventScroll: true });
      target.setSelectionRange(nextCaret, nextCaret);
    });
  };

  useEffect(() => {
    const el = inputRef.current;
    if (!el) return;
    el.focus({ preventScroll: true });
    if (caret === "all") {
      el.select(); // 双击：全选，便于直接替换
    } else {
      const end = el.value.length;
      el.setSelectionRange(end, end); // F2 / Enter / 直接输入：光标落在末尾
    }
  }, [caret]);

  const handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
    // 编辑中的按键一律不冒泡给网格：方向键 / Home / End 在 input 内移动光标，而不是移动选中格
    event.stopPropagation();
    if (doneRef.current) return;

    // Esc 的最高优先级：正在拾取引用时只取消这次拾取（编辑框不退出、文本不变）
    if (event.key === "Escape" && readPicking()) {
      event.preventDefault();
      onPickCancel();
      return;
    }

    // 候选面板打开时先接管 ↑↓ / Tab / Enter / Esc（与内容栏的优先级一致）
    if (suggest) {      const { items } = suggest.suggestion;
      const total = items.length;
      if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        event.preventDefault();
        const delta = event.key === "ArrowDown" ? 1 : -1;
        setSuggest((prev) => (prev ? { ...prev, index: (prev.index + delta + total) % total } : prev));
        return;
      }
      if (event.key === "Tab" || (event.key === "Enter" && !event.shiftKey)) {
        event.preventDefault();
        pickSuggestion(items[suggest.index]);
        return;
      }
      if (event.key === "Escape") {
        // 只关面板、不退出编辑；关掉后再按一次 Esc 才取消编辑（保持原有行为）
        event.preventDefault();
        dismissedRef.current = true;
        setSuggest(null);
        return;
      }
    }

    if (event.key === "Enter") {
      event.preventDefault();
      doneRef.current = true;
      onCommit(inputRef.current?.value ?? value, event.shiftKey ? "up" : "down");
      return;
    }
    if (event.key === "Tab") {
      event.preventDefault();
      doneRef.current = true;
      onCommit(inputRef.current?.value ?? value, event.shiftKey ? "left" : "right");
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      doneRef.current = true;
      onCancel(); // 丢弃本次输入，不回调 onCellCommit，选中格保持不变
    }
  };

  /** 失焦（点到别的单元格或页面上其它控件）即提交，选中格原地不动 */
  const handleBlur = () => {
    if (doneRef.current) return;
    doneRef.current = true;
    onCommit(inputRef.current?.value ?? value, "none");
  };

  return (
    <>
      <input
        ref={(el) => {
          inputRef.current = el;
          bindInput(el);
        }}
        value={value}
        onChange={(event) => setValue(event.target.value)}
        onKeyDown={handleKeyDown}
        onKeyUp={refreshSuggest}
        onClick={refreshSuggest}
        onSelect={refreshSuggest}
        onBlur={handleBlur}
        spellCheck={false}
        autoComplete="off"
        aria-label={`编辑 ${columnLabel(col)}${row + 1}`}
        className={`absolute top-0 z-[3] truncate border-r border-b border-line px-1.5 text-[12px] text-fg outline-none! ${
          // 正在编辑的命中格：底色换成命中提示色（编辑框本来是不透明的 bg-app，会把命中提示整个盖住）；
          // 当前命中还要把 accent 描边让给 2px 高对比描边，否则「停在编辑中的那一条」看不出来
          hitTint === "current"
            ? "bg-warning/40 ring-2 ring-inset ring-fg"
            : hitTint === "hit"
              ? "bg-warning/25 ring-2 ring-inset ring-accent"
              : "bg-app ring-2 ring-inset ring-accent"
        }`}
        // 可变行高：编辑框跟行一起长高，但 <input> 始终单行，文本按 ROW_H 行高在框内顶部对齐
        style={{ left, width, height, lineHeight: `${ROW_H}px` }}
      />
      {/* 候选面板挂到 body：内容区的 contain: layout paint 会让 fixed 定位错位并被裁剪 */}
      {suggest && typeof document !== "undefined"
        ? createPortal(<FormulaSuggestPanel state={suggest} onPick={pickSuggestion} />, document.body)
        : null}
    </>
  );
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
  /** 单元格按下（shiftKey 决定扩展选区还是重开一个 1×1 选区；clientX/Y 供拖拽会话起步） */
  onCellMouseDown: (row: number, col: number, shiftKey: boolean, clientX: number, clientY: number) => void;
  /** 行号栏按下：选中整行（Shift 扩展成多行），并进入整行拖拽 */
  onRowMouseDown: (row: number, shiftKey: boolean, clientX: number, clientY: number) => void;
  /** 单元格右键（clientX/Y 是弹菜单位置）；返回是否已由父组件接手（true 则拦掉系统菜单） */
  onCellContextMenu: (row: number, col: number, clientX: number, clientY: number) => boolean;
  /** 行号栏右键；返回是否已由父组件接手 */
  onRowContextMenu: (row: number, clientX: number, clientY: number) => boolean;
  /** 本行待提交的编辑（行级切片，见 pendingRowSlices）；undefined = 本行没有被改过的格 */
  pendingRow: RowPending | undefined;
  /** 本行正在编辑的列，-1 表示没有；只有它在 0..n 之间时才会多渲染一个 <input> */
  editingCol: number;
  /** 编辑框初始文本（仅 editingCol >= 0 时有意义；其余行固定传空串，保证 memo 的浅比较稳定） */
  editingText: string;
  editingCaret: EditCaret;
  /** 编辑接口：稳定引用，editable=false 时既不会渲染编辑框也不会接双击 */
  editor: RowEditorApi;
  /**
   * 选区在本行的**已渲染列窗口内**切片（含端点，数据列号）；-1 = 本行没有覆盖层。
   * 全部是原始值 props：拖拽时除了「端点真正变化的那几行」，其余行的 props 全等，
   * memo 的浅比较直接跳过（同一格内移动鼠标连一次渲染都不会有；
   * 横向拖出已渲染列窗口时也不会因为「真实端点变了」而白重渲染）。
   */
  selStartCol: number;
  selEndCol: number;
  /** 本行是否是选区的最上一行 / 最下一行（只有这两行画横向外边） */
  selTop: boolean;
  selBottom: boolean;
  /** 选区的真实左右边界是否落在已渲染列窗口内（不在就不画那条竖线） */
  selLeftEdge: boolean;
  selRightEdge: boolean;
  /**
   * 选区是否不止一格（>1）。Excel 语义：多格选区里的活动格（拖拽锚点）**不填充**，
   * 只留一个细边框；其余选中格才铺浅底色。单格/无选区时活动格保持原来的高亮样式不变。
   */
  multiSelection: boolean;
  /** 本行行高（可变行高；未开启换行时恒为 ROW_H，DOM 与改动前一致） */
  height: number;
  /**
   * 自动换行：开启后单元格不再 truncate，而是 pre-wrap + break-word 折行显示
   * （行高由 height 给出，正好等于折行数 × ROW_H）。关闭时 class 与 style 与改动前逐字节一致。
   */
  wrap: boolean;
  /** 公式拾取中：选区外框改用虚线，和普通选区区分开（只在拾取时为 true，DOM 默认不变） */
  picking: boolean;
  /** 行号栏下边界的拖动入口：调整本行行高（mousedown 只负责起手，移动/松手在 window 上） */
  onRowResizeStart: (row: number, clientX: number, clientY: number) => void;
  /** 双击行边界：本行按内容自动调整行高 */
  onRowResizeAuto: (row: number) => void;
  /**
   * 冻结列数（0 = 不冻结）：>0 时把「行号栏 + 前 N 列」放进一个行内 `sticky left-0` 块，
   * 水平滚动时整块贴住视口左边（与行号栏同一套机制，不另写第二套网格）。
   */
  frozenCols: number;
  /** 冻结块在内容坐标里的宽度（含行号栏宽度） */
  frozenWidth: number;
  /** 冻结列与滚动区之间的分隔线（只有冻结列后面还有滚动列时才画） */
  showColDivider: boolean;
  /** 本行是「表头样式」的冻结行（bg-panel / font-medium）：只有 headerRow 的那一行是 */
  panel: boolean;
  /** 冻结行与滚动区之间的分隔线（画在最后一个冻结行的下边缘） */
  showRowDivider: boolean;
  /** 查找命中的查表（`"row,col"`）；空集合表示这次渲染没有任何命中 */
  hitKeys: ReadonlySet<string>;
  /** 当前定位到的那条命中（`"row,col"`；空串 = 没有） */
  currentHitKey: string;
  /**
   * 本行是不是选区的最后一行（拖拽填充柄就挂在这一行的右下角）；-1 表示本行没有填充柄。
   * 值 = 选区右下角的列号（可能落在冻结列里 → 手柄要放进冻结块）。
   */
  fillHandleCol: number;
  /** 按下填充柄：起手一次拖拽填充会话（移动/松手在 window 上） */
  onFillHandleDown: (row: number, col: number, clientX: number, clientY: number) => void;
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
  onCellMouseDown,
  onRowMouseDown,
  onCellContextMenu,
  onRowContextMenu,
  pendingRow,
  editingCol,
  editingText,
  editingCaret,
  editor,
  selStartCol,
  selEndCol,
  selTop,
  selBottom,
  selLeftEdge,
  selRightEdge,
  multiSelection,
  height,
  wrap,
  picking,
  onRowResizeStart,
  onRowResizeAuto,
  frozenCols,
  frozenWidth,
  showColDivider,
  panel,
  showRowDivider,
  hitKeys,
  currentHitKey,
  fillHandleCol,
  onFillHandleDown,
}: SheetRowProps): JSX.Element {
  /**
   * 生成一格。冻结列与滚动列共用这一段（只是放进不同的容器），所以样式/事件/选区逻辑只有一处。
   * 注意：无选区时拼出的 class 与改造前逐字节一致（只读路径的硬要求）。
   */
  const buildCell = (col: number): JSX.Element => {
    const cell = getCell(row, col);
    const original = cell ? cell.v : "";
    // 待提交的编辑覆盖原始显示值；命中时右上角画一个小三角，让「哪些格被改过」一眼可见
    const pending = pendingRow?.get(col);
    const edited = pending !== undefined;
    const text = edited ? pending : original;
    const isActive = col === activeCol;
    // 本格是否落在选区（列区间已由父组件夹到已渲染列窗口）
    const inSelCol = selStartCol >= 0 && col >= selStartCol && col <= selEndCol;
    // 正在编辑的那一格：完全交给编辑框（input）渲染 —— 不再画活动格高亮、也不再铺选区浅底，
    // 否则 input 自己的 ring 会和活动格 ring/选区底色叠成「两层框」。
    const coveredByEditor = col === editingCol;
    // 多格选区里的活动格（拖拽锚点）：不铺底色，只留细边框（见 multiSelection 注释）
    const anchorInSelection = isActive && multiSelection && inSelCol && !coveredByEditor;
    const fill = inSelCol && !isActive && !coveredByEditor;
    /**
     * 查找命中提示：一层比选区更浅的底色（warning 系浅色，与蓝色的选区/活动格一眼可分）。
     * **命中一律有底色**（包括落在选区里的）—— 用户在「范围=选区」里查找时，命中全都在选区底色上，
     * 不给底色就等于「搜了看不见」。强度分档：
     * - 普通命中：`bg-warning/15`；选区里的命中：`bg-warning/40`（蓝色底上要更重才看得出来）
     * - **当前命中**（面板上的「当前/总数」）：`bg-warning/60` + **2px 中性高对比描边**（`ring-2 ring-fg`），
     *   一眼就能看到现在停在哪一条 —— 用户反馈过「当前命中只有一个很细的外框，不够明显」。
     *   描边用中性色而不是 warning 实色：浅色主题 warning 是深琥珀、深色主题是亮琥珀，
     *   铺在琥珀底色上会糊成一片；中性色在两种主题下都和底色、选区蓝拉开对比。
     */
    const cellKey = `${row},${col}`;
    const isHit = hitKeys.size > 0 && hitKeys.has(cellKey);
    const isCurrentHit = isHit && cellKey === currentHitKey;
    // 注意：无选区时下面拼出的 class 与改造前逐字节一致（只读路径的硬要求）；
    // 命中格的活动格底色让给命中色（同一个 background-color 只能有一个，靠 CSS 顺序决胜负太脆）；
    // 当前命中连活动格的 accent 描边也让位 —— 2px 高对比描边必须完整可见
    const activeClass = isActive && !coveredByEditor
      ? anchorInSelection
        ? ` z-[1]${isCurrentHit ? "" : " ring-1 ring-inset ring-accent"}`
        : ` z-[1]${isCurrentHit ? "" : " ring-2 ring-inset ring-accent"}${isHit ? "" : " bg-accent-soft"}`
      : "";
    const fillClass = fill && !isHit ? " bg-accent/10" : "";
    const hitInSelection = isHit && (fill || isActive);
    const hitClass = !isHit
      ? ""
      : isCurrentHit
        ? " bg-warning/60 ring-2 ring-inset ring-fg"
        : hitInSelection
          ? " bg-warning/40"
          : " bg-warning/15";
    // 冻结行的底色与浅填充 / 命中底色互斥（都是 background-color，留一个才不会看运气）：
    // 表头行用 bg-panel（与改动前一致），普通冻结行用 bg-app（不透明，挡住下面滚过去的行）
    const frozenClass = panel
      ? fill || isHit
        ? " font-medium"
        : " bg-panel font-medium"
      : frozen
        ? fill || isHit
          ? ""
          : " bg-app"
        : "";
    return (
      // 只画右边和下边：相邻单元格共用一条 1px 线，不会出现双线
      <div
        key={col} role="gridcell" aria-colindex={col + 1} aria-selected={isActive}
        title={edited ? `${original} → ${pending}` : text.length > 0 ? text : undefined}
        onMouseDown={(event) => {
          // 阻止浏览器默认的文本选择/焦点行为：拖拽框选才不会顺带选中文字
          event.preventDefault();
          onCellMouseDown(row, col, event.shiftKey, event.clientX, event.clientY);
        }}
        onContextMenu={(event) => {
          // 父组件接了右键回调才拦系统菜单，网格自己不渲染任何菜单
          if (onCellContextMenu(row, col, event.clientX, event.clientY)) event.preventDefault();
        }}
        onDoubleClick={editor.editable ? () => editor.startEdit(row, col) : undefined}
        // 换行关闭时保持原来的 truncate 单行样式（class/style 与改动前逐字节一致）；
        // 打开时改成 pre-wrap + break-word，高度由行高索引给出的 height 撑开（= 折行数 × ROW_H）。
        // overflow-hidden 是硬要求：折行只允许发生在**本单元格内部**，任何情况下都不能画到相邻行上。
        className={`absolute top-0 ${wrap ? "overflow-hidden whitespace-pre-wrap break-words" : "truncate"} border-r border-b border-line px-1.5 text-[12px] text-fg ${
          CELL_ALIGN[cell ? cell.t : "empty"]
        }${activeClass}${fillClass}${frozenClass}${hitClass}`}
        style={{ left: colOffsets[col], width: colWidths[col], height, lineHeight: `${ROW_H}px` }}
      >
        {text}
        {edited ? (
          // 已修改标记：右上角实心小三角（不参与命中与布局）
          <span
            aria-hidden="true"
            className="pointer-events-none absolute right-0 top-0 h-0 w-0 border-l-[6px] border-t-[6px] border-l-transparent border-t-accent"
          />
        ) : null}
      </div>
    );
  };

  /** 冻结列（0..frozenCols-1）与滚动列（firstCol..lastCol）分别渲染到两个容器里 */
  const frozenCells: JSX.Element[] = [];
  for (let col = 0; col < frozenCols; col += 1) frozenCells.push(buildCell(col));
  const cells: JSX.Element[] = [];
  for (let col = Math.max(firstCol, frozenCols); col <= lastCol; col += 1) cells.push(buildCell(col));

  const inSelection = selStartCol >= 0 && selEndCol >= selStartCol;
  /**
   * 选区外框在本行的分段。正在编辑的那一格「完全交给 input 渲染」，
   * 所以外框从它两侧断开（宁可在这里留一个缺口，也不要让两层框叠在一起）：
   * 分段内的上下边照画，左右竖边只由「真正包含选区左右边界的那一段」来画。
   */
  const frameSegments: Array<[number, number]> = [];
  if (inSelection) {
    const editCol = editingCol >= selStartCol && editingCol <= selEndCol ? editingCol : -1;
    if (editCol < 0) {
      frameSegments.push([selStartCol, selEndCol]);
    } else {
      if (selStartCol <= editCol - 1) frameSegments.push([selStartCol, editCol - 1]);
      if (editCol + 1 <= selEndCol) frameSegments.push([editCol + 1, selEndCol]);
    }
  }
  const gutterClass = `sticky left-0 border-b border-r border-line text-center text-[11px] ${
    activeCol >= 0 || inSelection ? "bg-accent-soft text-accent" : "bg-panel text-muted"
  } ${frozen ? "z-20" : "z-10"}`;
  /** 冻结块里的行号栏：块自己已经是 sticky left-0，栏内不需要再 sticky（改成绝对定位） */
  const frozenGutterClass = `absolute left-0 top-0 border-b border-r border-line text-center text-[11px] ${
    activeCol >= 0 || inSelection ? "bg-accent-soft text-accent" : "bg-panel text-muted"
  }`;
  const gutter = (
    <div
      role="rowheader"
      onMouseDown={(event) => {
        event.preventDefault();
        onRowMouseDown(row, event.shiftKey, event.clientX, event.clientY);
      }}
      onContextMenu={(event) => {
        if (onRowContextMenu(row, event.clientX, event.clientY)) event.preventDefault();
      }}
      className={frozenCols > 0 ? frozenGutterClass : gutterClass}
      style={{ width: GUTTER_W, height, lineHeight: `${ROW_H}px` }}
    >
      {label}
    </div>
  );
  /** 行高调整命中区：贴在行号栏下边界（只读网格同样可用）。stopPropagation 掉整行选中，
      所以拖动时不会顺手改选区；双击按内容自动调整行高 */
  const rowResizer = (
    <div
      role="separator"
      aria-label={`调整行高 ${row + 1}`}
      aria-orientation="horizontal"
      title="拖动调整行高，双击按内容自动调整"
      onMouseDown={(event) => {
        event.preventDefault();
        event.stopPropagation();
        onRowResizeStart(row, event.clientX, event.clientY);
      }}
      onDoubleClick={(event) => {
        event.stopPropagation();
        onRowResizeAuto(row);
      }}
      onContextMenu={(event) => event.preventDefault()}
      className="absolute left-0 z-30 cursor-row-resize hover:bg-accent/40"
      style={{ width: GUTTER_W, top: Math.max(0, height - RESIZE_HIT), height: RESIZE_HIT }}
    />
  );
  /**
   * 选区外框：只有边框，没有底色（浅底由每个单元格自己的 bg-accent/10 画，
   * 这样「多格选区里的活动格不填充」只需要不给那一格加 class，天然是一个洞）。
   * 绝对定位 + border-box + 写死宽高 → 完全不影响布局与行高；放在单元格之后所以边框压得住格线。
   * 列区间已经由父组件夹到「本行已渲染的列窗口」，被横向裁掉的那一侧不画竖线；
   * 编辑中的那一格会把本行的外框切成两段（见 frameSegments）。
   * 冻结列时按冻结边界把分段劈开：冻结那几段要放进冻结块里（否则会被滚走的单元格盖住）。
   */
  const buildFrame = (from: number, to: number): JSX.Element => (
    <div
      key={`${from}-${to}`}
      aria-hidden="true"
      className={`pointer-events-none absolute top-0 border-accent${picking ? " border-dashed" : ""}`}
      style={{
        left: colOffsets[from],
        width: Math.max(1, colOffsets[to] + colWidths[to] - colOffsets[from]),
        height,
        borderTopWidth: selTop ? 2 : 0,
        borderBottomWidth: selBottom ? 2 : 0,
        borderLeftWidth: selLeftEdge && from === selStartCol ? 2 : 0,
        borderRightWidth: selRightEdge && to === selEndCol ? 2 : 0,
      }}
    />
  );
  const frozenFrames: JSX.Element[] = [];
  const scrollFrames: JSX.Element[] = [];
  /** 本行实际渲染的滚动列从这一列开始（冻结列那几列在冻结块里，不在这里） */
  const scrollFirst = Math.max(firstCol, frozenCols);
  for (const [from, to] of frameSegments) {
    if (frozenCols > 0 && from < frozenCols) {
      frozenFrames.push(buildFrame(from, Math.min(to, frozenCols - 1)));
      // 冻结列与滚动列之间可能有「被滚过去的列」没渲染：滚动那一段夹到实际渲染的第一列
      if (to >= scrollFirst) scrollFrames.push(buildFrame(scrollFirst, to));
    } else if (to >= scrollFirst) {
      scrollFrames.push(buildFrame(Math.max(from, scrollFirst), to));
    }
  }
  /** 编辑框：编辑冻结列里的格子时要放进冻结块（否则会被滚走的单元格盖住 / 跟着滚走） */
  const editorEl = editor.editable && editingCol >= 0 ? (
    <SheetCellEditor
      key={`${row}:${editingCol}`}
      row={row}
      col={editingCol}
      left={colOffsets[editingCol]}
      width={colWidths[editingCol]}
      height={height}
      text={editingText}
      caret={editingCaret}
      bindInput={editor.bindInput}
      onCommit={editor.commitEdit}
      onCancel={editor.cancelEdit}
      readCompleter={editor.readCompleter}
      bindApi={editor.bindApi}
      readPicking={editor.readPicking}
      onPickCancel={editor.cancelPick}
      onFormulaModeChange={editor.notifyFormulaMode}
      hitTint={
        hitKeys.size > 0 && hitKeys.has(`${row},${editingCol}`)
          ? `${row},${editingCol}` === currentHitKey
            ? "current"
            : "hit"
          : "none"
      }
    />
  ) : null;
  const editorInFrozen = editorEl !== null && frozenCols > 0 && editingCol < frozenCols;
  /**
   * 拖拽填充柄（Excel 的小方块）：贴在选区右下角的格子角上，鼠标移上去是 crosshair。
   * 用中性深色方块（Excel 同款观感）：它既不能和琥珀色命中混淆，也不能和蓝色选区混淆。
   * 手柄放在**本行内部**（而不是独立图层）：冻结行会跟着 sticky、冻结列会跟着冻结块走，
   * 滚动、换行、虚拟化都不需要额外换算。列落在冻结列里时放进冻结块（与编辑框同一套处理）。
   */
  const fillHandle =
    fillHandleCol >= 0 ? (
      <div
        role="presentation"
        aria-hidden="true"
        data-fill-handle="true"
        title="拖动以按规律填充（Esc 取消）"
        onMouseDown={(event) => {
          // 不触发选区拖拽 / 不进入编辑：手柄自己的会话完全独立
          event.preventDefault();
          event.stopPropagation();
          onFillHandleDown(row, fillHandleCol, event.clientX, event.clientY);
        }}
        onDoubleClick={(event) => event.stopPropagation()}
        onContextMenu={(event) => {
          event.preventDefault();
          event.stopPropagation();
        }}
        className="absolute z-[4] cursor-crosshair border border-app bg-fg"
        style={{
          left: Math.max(0, colOffsets[fillHandleCol] + colWidths[fillHandleCol] - FILL_HANDLE - 1),
          top: Math.max(0, height - FILL_HANDLE - 1),
          width: FILL_HANDLE,
          height: FILL_HANDLE,
        }}
      />
    ) : null;
  const fillHandleInFrozen = fillHandle !== null && frozenCols > 0 && fillHandleCol < frozenCols;
  /** 冻结行与滚动区之间的分隔线：比普通网格线略明显（2px 强调色），只画在最后一个冻结行下边缘 */
  const rowDivider = showRowDivider ? (
    <div
      aria-hidden="true"
      data-freeze-divider="row"
      className="pointer-events-none absolute bottom-0 left-0 z-[2] h-[2px] bg-accent/40"
      style={{ width }}
    />
  ) : null;
  return (
    <div
      role="row"
      className={
        frozen
          ? `sticky z-[15] border-b border-line ${panel ? "bg-panel" : "bg-app"}`
          : "absolute left-0"
      }
      // 开启换行时给整行加一道 paint 裁剪：即使某个单元格的高度算错，内容也只会被裁在行内，
      // 不会画到下一行去。用 contain 而不是 overflow:hidden —— 后者会把行变成滚动容器，
      // 行号栏的 sticky left-0 就会改成相对这一行定位，横向滚动时不再吸附。
      style={wrap ? { top, width, height, contain: "paint" } : { top, width, height }}
    >
      {frozenCols > 0 ? (
        /*
         * 冻结块：行号栏 + 前 N 列 + 它们的选区外框 / 编辑框。
         * sticky left-0 让它水平滚动时贴住视口左边（与行号栏同一套机制），块内仍用内容坐标定位，
         * 所以选区、编辑框、拖拽换算都不需要额外偏移。宽度 = 冻结区宽度，不会盖住滚动列。
         *
         * **必须有不透明底色**：块压在滚动列之上，冻结列的单元格本身是透明的，
         * 少了这层底色就会「下面的内容透过首列显示出来」，看起来像第一列没被置顶（用户反馈）。
         * 底色跟本行的冻结行底色一致（表头行 bg-panel，其余 bg-app），否则表头行的冻结列会串色。
         */
        <div className={`sticky left-0 z-10 ${panel ? "bg-panel" : "bg-app"}`} style={{ width: frozenWidth, height }}>
          {gutter}
          {rowResizer}
          {frozenCells}
          {frozenFrames}
          {editorInFrozen ? editorEl : null}
          {fillHandleInFrozen ? fillHandle : null}
          {showColDivider ? (
            <div
              aria-hidden="true"
              data-freeze-divider="col"
              className="pointer-events-none absolute top-0 h-full w-[2px] bg-accent/40"
              style={{ left: Math.max(0, frozenWidth - 2) }}
            />
          ) : null}
        </div>
      ) : (
        <>
          {gutter}
          {rowResizer}
        </>
      )}
      {cells}
      {scrollFrames}
      {editorInFrozen ? null : editorEl}
      {fillHandleInFrozen ? null : fillHandle}
      {rowDivider}
    </div>
  );
});

/* -------------------------------- 主组件 -------------------------------- */

/** 空命中集合：不传 findHits 时所有行共用它，保证 memo 的 props 引用稳定 */
const EMPTY_HIT_KEYS: ReadonlySet<string> = new Set<string>();

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
  editable = false,
  pendingCells,
  onCellCommit,
  selection,
  onSelectionChange,
  onSelectionCommit,
  onClearSelection,
  onCopySelection,
  onPaste,
  onCellContextMenu,
  completeFormula,
  autoFitToken,
  wrapText = false,
  onLayoutChange,
  pickReference = false,
  onPickReference,
  formulaRanges,
  columnWidths,
  rowHeights,
  onColumnResize,
  onRowResize,
  autoFitRowsToken,
  onRowsResize,
  onColumnsResize,
  freezeRows = 0,
  freezeCols = 0,
  findHits,
  currentHit,
  scrollTarget,
  onFillRange,
}: SheetGridProps): JSX.Element {
  const scrollerRef = useRef<HTMLDivElement>(null);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [viewport, setViewport] = useState({ w: 0, h: 0 });
  const [active, setActive] = useState<{ row: number; col: number } | null>(activeCell ?? null);
  /** 正在编辑的单元格（null = 没在编辑）。它只影响 memo 命中的那一行重渲染 */
  const [editing, setEditing] = useState<EditingState | null>(null);
  /**
   * 选区（渲染用的矩形）。语义与 activeCell 一致：
   * 非受控（selection 传 undefined）时完全由网格维护；父组件传具体值/ null 时按受控处理。
   * 点选单格也**不会**变成 null，而是收缩成 1×1（父组件可以只用一个 range 就表达统计口径）。
   * 初值直接取自 props（而不是等 effect 回填），这样首帧/SSR 就与父组件的选区一致，不会闪一下。
   */
  const [range, setRange] = useState<SheetRange | null>(() => {
    if (selection !== undefined) return clampRange(selection, rows, cols);
    return activeCell ? singleRange(activeCell.row, activeCell.col) : null;
  });
  const lastViewportKeyRef = useRef("");

  // 外部回调与数据放进 ref：事件回调因此保持引用稳定，父组件重渲染（含 pendingCells 换新 Map）
  // 都不会导致所有行重新渲染
  const callbacksRef = useRef({
    onActiveCell,
    onCellCommit,
    getCell,
    pendingCells,
    editable,
    onSelectionChange,
    onSelectionCommit,
    onClearSelection,
    onCopySelection,
    onPaste,
    onCellContextMenu,
    completeFormula,
    onLayoutChange,
    onPickReference,
    onColumnResize,
    onRowResize,
    onRowsResize,
    onColumnsResize,
    onFillRange,
    /** selection 是否受控（父组件传了具体值或 null） */
    selectionControlled: selection !== undefined,
    /** 列宽 / 行高是否受控（父组件传了 columnWidths / rowHeights，见 applyColumnWidth 注释） */
    controlledColumns: columnWidths !== undefined,
    controlledRows: rowHeights !== undefined,
  });
  callbacksRef.current = {
    onActiveCell,
    onCellCommit,
    getCell,
    pendingCells,
    editable,
    onSelectionChange,
    onSelectionCommit,
    onClearSelection,
    onCopySelection,
    onPaste,
    onCellContextMenu,
    completeFormula,
    onLayoutChange,
    onPickReference,
    onColumnResize,
    onRowResize,
    onRowsResize,
    onColumnsResize,
    onFillRange,
    selectionControlled: selection !== undefined,
    controlledColumns: columnWidths !== undefined,
    controlledRows: rowHeights !== undefined,
  };

  // 编辑态的「同步值」：state 更新是异步的，而 Enter 之后紧跟的 blur、滚动提交等
  // 都要在同一轮事件里判断「是否还在编辑」，因此用 ref 作为权威副本
  const editingRef = useRef<EditingState | null>(null);
  /** 当前编辑框的真实 input 节点：滚动提交时直接读它的当前值 */
  const editorInputRef = useRef<HTMLInputElement | null>(null);
  /** finishEditing 的 ref 副本：handleScroll 的依赖数组必须保持为空（引用稳定），不能直接闭包引用它 */
  const finishEditingRef = useRef<(text: string | undefined, move: CommitMove) => void>(() => {});
  /** active 的同步副本：只读 ref 的编辑回调靠它判断「这一格是不是已经选中」，避免把 active 放进依赖导致回调反复重建 */
  const activeRef = useRef(active);
  activeRef.current = active;
  /** Shift 扩展的锚点（选区固定角）；点选/键盘移动时跟着走，Shift 操作时保持不动 */
  const anchorRef = useRef<{ row: number; col: number } | null>(activeCell ? { row: activeCell.row, col: activeCell.col } : null);
  /** 正在进行的拖拽会话（null = 没有在拖拽）；window 监听器与自动滚动都靠它判断 */
  const dragRef = useRef<DragSession | null>(null);
  /** 公式拾取模式：编辑框里的文本以 = 开头时为 true（由编辑框通知，只在真的翻转时 setState）。 */
  const formulaModeRef = useRef(false);
  /**
   * 外部拾取是否可用：内容栏正在编辑公式且当前不是只读网格（只读下 pickReference 无效，不回调）。
   */
  const externalPick = pickReference && !readOnly;
  /** 拾取会话开始前的选区（Esc 取消拾取时恢复它） */
  const pickStartRef = useRef<SheetRange | null>(null);
  /** 编辑框注册进来的改写能力（拾取结束要把引用写回去） */
  const editorApiRef = useRef<CellEditorApi | null>(null);
  /** 拾取中的视觉（虚线外框）：只在拾取开始/结束时变化 */
  const [picking, setPicking] = useState(false);
  /**
   * 手动拖出来的列宽（列 → px）。**优先于采样值与「自动调整列宽」的结果，且只在会话内有效**
   * （暂不写回 xlsx：持久化是后续单独一件事）。
   */
  const manualColsRef = useRef(new Map<number, number>());
  /** 手动拖过行高的行（这些行不再被换行扫描改写，双击边界即可恢复自动） */
  const manualRowsRef = useRef(new Set<number>());
  /** 列宽地图的版本号：手动拖列宽时 +1（行高那边复用 heightsVersion） */
  const [colsVersion, setColsVersion] = useState(0);
  /** 正在进行的宽高调整会话（null = 没有在调整） */
  const resizeRef = useRef<{
    kind: "col" | "row";
    index: number;
    startClient: number;
    startSize: number;
    size: number;
    min: number;
    max: number;
    /** 会话开始时的容器左上角（client 坐标）：指示线要换算成容器内部坐标 */
    originLeft: number;
    originTop: number;
  } | null>(null);
  /** 拖动时的指示线（client 坐标；null = 没在拖）：拖动过程中给用户一条可见的落点线 */
  const [resizeGuide, setResizeGuide] = useState<{ kind: "col" | "row"; client: number } | null>(null);

  const empty = rows === 0 || cols === 0;
  /**
   * 冻结的行数：`freezeRows` 与 `headerRow` 是同一套实现 —— headerRow 就是「冻结 1 行 + 表头样式」，
   * 两者同时传时取较大值（**不会**出现两层表头）；都为 0 时行为与改动前逐字节一致。
   */
  const explicitFreezeRows = Math.max(0, Math.floor(freezeRows) || 0);
  const freezeRowCount = Math.max(explicitFreezeRows, Math.min(rows, headerRow ? 1 : 0), 0);
  /** 表头样式的行（只有 headerRow 的第 0 行有 bg-panel / font-medium） */
  const headerStyledRow = headerRow && freezeRowCount > 0 ? 0 : -1;
  const headerOffset = freezeRowCount;
  /** 自动换行是否生效：不传 / false 时整条换行链路（扫描、行高、折行 class）完全空转，DOM 与改动前一致 */
  const wrapEnabled = Boolean(wrapText);
  /** 参与滚动的行数：冻结行不再占用滚动高度 */
  const scrollRows = Math.max(0, rows - headerOffset);

  /** 渲染与回调统一使用「夹到数据范围内」的选区；空表为 null（值未变时引用稳定） */
  const sel = useMemo(() => clampRange(range, rows, cols), [range, rows, cols]);
  /** 选区的同步副本：键盘 / window 事件回调靠它读最新值，省得把 sel 放进依赖导致回调重建 */
  const rangeRef = useRef<SheetRange | null>(sel);
  rangeRef.current = sel;

  /**
   * 已计算过的「非默认行高」（键 = 数据行号）。只对**已加载**的行计算折行高度：
   * 未加载的行拿不到内容，一律先用 ROW_H，等数据到达（getCell 换引用）后再修正。
   */
  const heightsRef = useRef(new Map<number, number>());
  /** 行高地图的版本号：只有它变化才重建行高索引（滚动本身绝不重建） */
  const [heightsVersion, setHeightsVersion] = useState(0);
  /** 布局版本号：行高或列宽真的变了就 +1，用来回调 onLayoutChange */
  const [layoutVersion, setLayoutVersion] = useState(0);

  /** 自动调整列宽的结果（null = 还没调过，用默认采样列宽） */
  const [fitWidths, setFitWidths] = useState<number[] | null>(null);

  // 列宽只在列数 / 总行数 / getCell 引用（数据身份）变化时重算，滚动时不重算
  const sampledWidths = useMemo(
    () => computeColWidths(cols, Math.min(rows, SAMPLE_ROWS), getCell),
    [cols, rows, getCell],
  );
  /**
   * 最终列宽，优先级：**受控 columnWidths（父组件持有，可撤销） > 内部手动拖过的列
   * > 自动调整列宽的采样结果 > 默认采样**。
   * 受控模式下内部不再记手动列宽（否则会和 props 打架），且受控值放在自动调整结果之后合并，
   * 所以「自动调整列宽」永远不会覆盖用户拖出来 / 撤销恢复回来的宽度。
   */
  const colWidths = useMemo(() => {
    const base = fitWidths && fitWidths.length === cols ? fitWidths : sampledWidths;
    const controlled = controlledSizes(columnWidths);
    // 受控模式下内部那份手动列宽完全不参与（否则会和 props 打架，撤销也撤不干净）
    const manual = controlled === null ? manualColsRef.current : null;
    const hasManual = manual !== null && manual.size > 0;
    const hasControlled = controlled !== null && controlled.size > 0;
    if (!hasManual && !hasControlled) return base;
    const merged = base.slice();
    if (hasManual && manual !== null) {
      for (const [col, width] of manual) if (col >= 0 && col < cols) merged[col] = width;
    }
    if (hasControlled && controlled !== null) {
      for (const [col, width] of controlled) {
        const next = clampManualColWidth(width);
        if (next !== null && col >= 0 && col < cols) merged[col] = next;
      }
    }
    return merged;
    // colsVersion 代表 manualColsRef 的内容（每次改地图都会 +1）
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fitWidths, sampledWidths, cols, colsVersion, columnWidths]);

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

  /**
   * 有效行高表：内部表（换行折出来的行高 + 未受控时手动拖过的行高）叠加**受控 rowHeights**。
   * 受控值优先（父组件持有它，撤销就是把旧对象传回来）；受控模式下内部表里同名的行直接让位，
   * 所以「拖完 → 回调 → 父组件把值传回来」这条链路只有一个真相来源。
   */
  const effectiveRowHeights = useMemo(() => {
    const controlled = controlledSizes(rowHeights);
    const local = heightsRef.current;
    if (controlled === null || controlled.size === 0) return local;
    const merged = new Map(local);
    for (const [row, height] of controlled) {
      const next = clampManualRowHeight(height);
      if (next === null) continue;
      merged.set(row, next);
    }
    return merged;
    // heightsRef 的内容由 heightsVersion 代表（每次改地图都会 +1）
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rowHeights, heightsVersion]);

  /**
   * 行高索引：scrollTop ↔ 行号、每行顶部 y 与内容总高度的**唯一来源**。
   * 表里只有「高度 ≠ ROW_H 的行」：换行折出来的行，以及**用户手动拖过 / 受控覆盖的行高**。
   * 表为空时索引退化成乘法，DOM 与原实现逐字节一致。
   * 依赖里刻意不放滚动位置：滚动不重建索引，这是百万行不退化的关键。
   */
  const rowIndex = useMemo(
    () => buildRowIndex(rows, headerOffset, effectiveRowHeights),
    // 未受控时 effectiveRowHeights 就是 heightsRef.current（引用恒定），所以 heightsVersion 必须留在依赖里；
    // 受控时它的引用跟着 rowHeights 变，两个依赖缺一不可。
    [rows, headerOffset, effectiveRowHeights, heightsVersion],
  );

  /**
   * 行高是否**全表都是默认值**（= 行高索引没有例外行）。
   * 用来决定「未渲染区域的占位行线」能不能画：那层是固定 24px 间距的 repeating-linear-gradient，
   * 一旦有行高 ≠ ROW_H（换行折行出来的高行、手动拖过的行高），间距就与真实行边界脱节 ——
   * 用户反馈的「调整行高后整个网格布满横线」正是这个。宁可少画线，也不画错位置的线：
   * 存在非默认行高时这一层直接不画，行边界交给每一行自己的 border-b（画在真实行高处）。
   * totalHeight 只在没有例外行时才等于 行数 × ROW_H（所有行高都 ≥ ROW_H），所以这是等价判据。
   */
  const uniformRowHeight = rowIndex.totalHeight === scrollRows * ROW_H;

  /**
   * 冻结列数（0 = 不冻结）。冻结列渲染在每个行的**行内 sticky 块**里（块内是行号栏 + 冻结列，
   * 水平滚动时整块贴住视口左边），因此不需要第二套网格；不冻结时整条分支完全不渲染，
   * DOM 与改动前逐字节一致。
   */
  const freezeColCount = Math.max(0, Math.min(cols, Math.floor(freezeCols) || 0));
  /** 冻结区（行号栏 + 冻结列）在内容坐标里的右边界 */
  const frozenWidth = freezeColCount > 0 ? colOffsets[freezeColCount - 1] + colWidths[freezeColCount - 1] : GUTTER_W;
  /**
   * 冻结列的列号列表。**不能**用 visibleCols 过滤出来：visibleCols 是「滚动列窗口」（从 viewRange.firstCol
   * 起，而它已经 ≥ freezeColCount），拿它去筛冻结列永远是空的 —— 冻结列的列标与列宽命中区会整块消失。
   */
  const frozenColList = useMemo(
    () => Array.from({ length: freezeColCount }, (_, col) => col),
    [freezeColCount],
  );

  /**
   * 查找命中的查表（`"row,col"` → 命中）。**不新增任何 DOM 节点**：命中只体现为格子上的一个 class，
   * 因此天然只作用于「渲染窗口内」的格子（500 条命中也不会塞 500 个节点），
   * 冻结行/冻结列里的命中同样由同一个 buildCell 渲染。不传命中时复用同一个空集合（引用稳定）。
   */
  const hitKeys = useMemo(() => {
    if (!findHits || findHits.length === 0) return EMPTY_HIT_KEYS;
    const set = new Set<string>();
    for (const hit of findHits) set.add(`${hit.row},${hit.col}`);
    return set;
  }, [findHits]);
  /** 当前定位到的命中（面板上的「当前/总数」）：空串表示没有 */
  const currentHitKey = currentHit ? `${currentHit.row},${currentHit.col}` : "";
  /** 冻结行/列与滚动区之间的分隔线：只有**显式**冻结了才画（headerRow 单独使用时 DOM 与改动前一致） */
  const showRowDivider = explicitFreezeRows > 0 && scrollRows > 0;
  const showColDivider = freezeColCount > 0 && freezeColCount < cols;

  /**
   * 顶部冻结区总高度（列字母表头 + 冻结行的高度之和）。
   * 逐行累加而不是「行数 × 第 0 行行高」：多行冻结时各行高度可能不同（换行折行、手动调整）。
   */
  const stickyH = useMemo(() => {
    let total = HEADER_H;
    for (let row = 0; row < headerOffset; row += 1) total += rowIndex.heightOf(row);
    return total;
  }, [rowIndex, headerOffset]);

  /** 指针 → 行列换算需要的几何量（window 监听器只挂一次，因此用 ref 读最新值） */
  const geomRef = useRef({
    colOffsets, colWidths, rows, cols, stickyH, headerOffset, scrollRows, rowIndex,
    frozenWidth, freezeColCount,
  });
  geomRef.current = {
    colOffsets, colWidths, rows, cols, stickyH, headerOffset, scrollRows, rowIndex,
    frozenWidth, freezeColCount,
  };

  // 内容宽度至少铺满视口，保证表头背景与右侧空白区域完整
  const contentWidth = Math.max(cols > 0 ? colOffsets[cols - 1] + colWidths[cols - 1] : GUTTER_W, viewport.w);

  /**
   * 滚动位置用 requestAnimationFrame 合并：
   * 拖动滚动条时 scroll 事件可能一帧触发多次，逐次 setState 会让 React 落后于滚动位置，
   * 表现为「先白屏、再补上数据」。合并后每帧最多渲染一次，且总是按最新位置渲染。
   */
  const pendingScrollRef = useRef<{ top: number; left: number } | null>(null);
  const rafRef = useRef<number | null>(null);
  /**
   * 程序化修正 scrollTop（行高变化后的锚点保持）会触发 scroll 事件；
   * 用它把这**不是用户滚动**的那一次区分出来，避免顺手把正在编辑的单元格提交掉。
   */
  const programmaticScrollRef = useRef(false);

  const handleScroll = useCallback(() => {
    const el = scrollerRef.current;
    if (!el) return;
    // 编辑中滚动（滚轮 / 拖动滚动条）：先把当前输入提交掉（原地，不移动选中格），
    // 否则编辑框会跟着内容一起滚走，视觉上与「正在编辑这一格」脱节
    if (editingRef.current && !programmaticScrollRef.current) finishEditingRef.current(undefined, "none");
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

  /**
   * 去掉「点一下单元格，网格外圈就多出一圈」。
   *
   * 成因：styles.css 里的全局 `*:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px }`
   * 是**未分层**规则，而 Tailwind 的 `.outline-none` 在 `@layer utilities` 里 ——
   * 未分层规则优先级更高，所以 class 上的 outline-none 根本压不住它；
   * 而单元格 mousedown 会 `scrollerRef.current.focus()`（见 handleCellMouseDown），
   * 于是整个网格被描上一圈 2px 强调色焦点框（编辑框上还会再叠一圈）。
   *
   * 为什么用行内样式而不是把 class 改成 `outline-none!`：
   * 行内样式在作者级里优先级最高，能稳定压过未分层规则；同时它不参与 SSR，
   * 只读路径也不会走到这里 —— test.mjs / edit.mjs 对「未编辑时的 DOM 逐字节一致」有硬断言。
   * 键盘导航的焦点提示由活动格 ring + 行号/列标高亮负责，不需要这圈 outline。
   */
  useLayoutEffect(() => {
    const el = scrollerRef.current;
    if (!el || !editable) return;
    el.style.outline = "none";
  }, [editable]);

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
    // 虚拟化核心：行高索引把 scrollTop 换算成第一个可见的虚拟行（未开启换行时就是 scrollTop / ROW_H）；
    // 可见行数由「视口高度 - 顶部冻结区高度」决定，最后上下各加 OVERSCAN 行缓冲。
    const firstVisible = rowIndex.virtualAt(scroll.top);
    const lastVisible = Math.max(firstVisible, rowIndex.virtualAt(scroll.top + Math.max(0, viewport.h - stickyH)));
    const firstRow = Math.max(0, firstVisible - OVERSCAN);
    const lastRow = Math.min(scrollRows - 1, lastVisible + OVERSCAN);

    // 列窗口：横向滚动时左侧有行号栏（+ 冻结列）遮住，可见区左边界要加上它们
    const left = scroll.left + frozenWidth;
    const right = scroll.left + Math.max(viewport.w, GUTTER_W);
    let firstCol = freezeColCount;
    while (firstCol < cols - 1 && colOffsets[firstCol] + colWidths[firstCol] <= left) firstCol += 1;
    let lastCol = cols - 1;
    while (lastCol > 0 && colOffsets[lastCol] >= right) lastCol -= 1;

    return {
      firstRow,
      lastRow,
      // 冻结列不参与滚动窗口：它们每个行都会渲染（见 SheetRow 的冻结块），这里从 freezeColCount 起
      firstCol: Math.max(freezeColCount, firstCol - COL_OVERSCAN),
      lastCol: Math.min(cols - 1, lastCol + COL_OVERSCAN),
      firstVisible,
      lastVisible,
    };
  }, [empty, scroll.top, scroll.left, viewport.w, viewport.h, stickyH, scrollRows, cols, colOffsets, colWidths, rowIndex, frozenWidth, freezeColCount]);

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

  /* ------------------ 自动换行：行高扫描 / 锚点保持 / 列宽自适应 ------------------ */

  /**
   * 已按「哪个数据版本 + 哪套列宽」扫过行高的行。
   * 只有数据变了（getCell 换引用）、待提交编辑变了（pendingCells 换引用）或列宽变了（换行折行点随之改变）
   * 才需要重扫；否则滚回同一批行时直接跳过内容扫描 —— 滚动开销与未开启换行时相同。
   */
  const scannedRef = useRef<{
    getCell: GetCell;
    pendingCells: ReadonlyMap<string, string> | undefined;
    colWidths: number[];
    rows: Map<number, true>;
    /** 上一次扫描时「行高由 props 受控」的行：控制权一变就要重新测量（撤销后不能留着旧值） */
    controlled: Map<number, true>;
  } | null>(null);
  /** 折行测量用的字体（从真实单元格上读一次后缓存） */
  const measureFontRef = useRef("");

  /** 受控行高表（rowHeights → Map）：换行扫描要跳过这些行，并在控制权变化时让它们重新测量 */
  const controlledRowHeights = useMemo(() => controlledSizes(rowHeights), [rowHeights]);

  /** 折行测量字体：优先用单元格的**真实字体**（canvas measureText 才准），拿不到就退回 12px sans-serif */
  const resolveMeasureFont = useCallback((): string => {
    if (!measureFontRef.current) {
      const sample = scrollerRef.current?.querySelector('[role="gridcell"]');
      const font = sample && typeof window !== "undefined" ? window.getComputedStyle(sample).font : "";
      measureFontRef.current = font && font.trim().length > 0 ? font : "12px sans-serif";
    }
    return measureFontRef.current;
  }, []);

  /**
   * 当前可视窗口的 ref 副本：双击列边界的采样要用「此刻屏幕上看到的那几行」，
   * 但这个值不能进 useCallback 依赖（滚动一次就重建回调，会让所有行白重渲染），所以走 ref。
   */
  const viewRangeRef = useRef(viewRange);
  viewRangeRef.current = viewRange;

  /**
   * pendingCells 里出现过待提交值的行（key 形如 `行,列`）。
   * 「这一行有没有可测内容」要连**刚输入还没保存**的格子一起算：用户往文件里本来没有数据的空格子
   * 输入长内容时 getCell 是 undefined（文件里没有这个单元格），内容只在 pendingCells 里 ——
   * 这种行同样要能按内容撑高、也要能参与一键自动调整行高。
   */
  const pendingRowSet = useMemo(() => {
    const set = new Set<number>();
    if (pendingCells) {
      for (const key of pendingCells.keys()) {
        const sep = key.indexOf(",");
        if (sep <= 0) continue;
        const row = Number(key.slice(0, sep));
        if (Number.isInteger(row) && row >= 0) set.add(row);
      }
    }
    return set;
  }, [pendingCells]);
  /**
   * 上面那张表的 ref 副本：measureRow 必须**引用稳定**（它会经 autoFitRow → SheetRow 的 props 传下去，
   * 身份一变所有行都会重渲染），所以走 ref 读最新值，而不是把它写进 useCallback 依赖。
   * 数据变化照样会重扫：扫描缓存的失效判据里本来就有 pendingCells 的引用比较。
   */
  const pendingRowSetRef = useRef(pendingRowSet);
  pendingRowSetRef.current = pendingRowSet;

  /**
   * 按内容算一行的行高，并顺带告诉调用方这一行**有没有可测内容**。
   * 文本用 readCellText 取 —— **待提交的编辑优先**，所以刚输入的长串也会把行撑高。
   * 关闭换行时行高恒为当前行高（不折行）：内容只有一行（truncate 截断），
   * 所以双击行边界「按内容自动调整」也只能得到 ROW_H，不能按折行数把行撑成几十行高。
   *
   * 「有可测内容」= 窗口列范围内拿到过 getCell 对象，**或** pendingCells 里有这一行的待提交值。
   * 行高扫描与一键自动调整行高共用这一个入口（可测性判定只有这一处，避免两边各写一份而漏掉待提交编辑）。
   */
  const measureRow = useCallback(
    (row: number): { height: number; hasContent: boolean } => {
      const live = callbacksRef.current;
      const scanCols = Math.min(cols, WRAP_SCAN_COLS);
      let hasContent = pendingRowSetRef.current.has(row);
      let lines = 1;
      const font = wrapEnabled ? resolveMeasureFont() : "";
      for (let col = 0; col < scanCols; col += 1) {
        const cell = live.getCell(row, col);
        if (cell !== undefined) hasContent = true;
        if (!wrapEnabled) continue; // 关闭换行：只需要知道有没有内容，不必逐格折行测量
        const text = live.pendingCells?.get(`${row},${col}`) ?? cell?.v ?? "";
        if (text.length === 0) continue;
        const count = wrapLineCount(text, colWidths[col] - CELL_PAD_X, font);
        if (count > lines) lines = count;
        if (lines >= WRAP_MAX_LINES) break;
      }
      return { height: wrapEnabled ? lines * ROW_H : ROW_H, hasContent };
    },
    [wrapEnabled, colWidths, cols, resolveMeasureFont],
  );

  /** 只要行高时的便捷入口（双击行边界、拖拽会话用） */
  const measureRowHeight = useCallback((row: number): number => measureRow(row).height, [measureRow]);

  /**
   * 换行高度扫描：只扫「渲染窗口 + 一小段余量」里的行（未加载的行 getCell 返回 undefined，
   * 自然维持默认行高，等数据到达后再修正）。算出的高度写进 heightsRef，
   * 真的变了才 bump 版本号 —— 这一步在 useLayoutEffect 里做，React 会在 paint 之前同步重渲染，
   * 因此不会出现「先按默认行高画一帧、再跳一下」。
   *
   * 依赖里带上 getCell / pendingCells 是必需的：数据窗口到达（父组件换 getCell）或待提交编辑变化时，
   * 之前算过的行高可能已经过期，必须重扫一次；里面的 token 比较会因此清空扫描缓存。
   * **手动拖过行高的行直接跳过**（用户手动优先，直到双击边界恢复自动）。
   */
  useLayoutEffect(() => {
    if (!wrapEnabled || empty) return;
    const live = callbacksRef.current;
    const cache = scannedRef.current;
    if (
      !cache ||
      cache.getCell !== live.getCell ||
      cache.pendingCells !== live.pendingCells ||
      cache.colWidths !== colWidths
    ) {
      scannedRef.current = {
        getCell: live.getCell,
        pendingCells: live.pendingCells,
        colWidths,
        rows: new Map(),
        controlled: new Map(),
      };
    }
    const scanned = scannedRef.current?.rows;
    const scannedControlled = scannedRef.current?.controlled;
    if (!scanned || !scannedControlled) return;
    const window_ = heightScanWindow(viewRange, headerOffset, rows);
    const first = window_.first;
    const last = window_.last;
    let changed = false;
    for (let row = first; row <= last; row += 1) {
      const isControlled = controlledRowHeights !== null && controlledRowHeights.has(row);
      if (scanned.has(row)) {
        // 行高的「控制权」没变就跳过（滚动时的主要开销都在这里省掉）
        if ((scannedControlled.get(row) === true) === isControlled) continue;
        // 控制权变了（受控 ↔ 自动，例如撤销恢复）：让这一行重新测量，不能留着旧值
        scanned.delete(row);
        scannedControlled.delete(row);
      }
      scanned.set(row, true);
      if (isControlled) {
        // 受控行高优先：内部表里同名的行让位（撤销后由扫描按当前列宽/内容重算）
        scannedControlled.set(row, true);
        if (heightsRef.current.delete(row)) changed = true;
        continue;
      }
      if (manualRowsRef.current.has(row)) continue; // 手动行高优先
      // 用与一键自动调整行高同一个入口测量（待提交编辑同样算「有内容」，纯待提交的行也会被撑高）；
      // 没有内容的行也照走一遍 —— 它测出来就是 ROW_H，会把之前留下的高行高收回去。
      const next = measureRow(row).height;
      const prev = heightsRef.current.get(row);
      if (next === prev || (next === ROW_H && prev === undefined)) continue;
      if (next === ROW_H) heightsRef.current.delete(row);
      else heightsRef.current.set(row, next);
      changed = true;
    }
    if (changed) {
      setHeightsVersion((value) => value + 1);
      setLayoutVersion((value) => value + 1);
    }
  }, [wrapEnabled, empty, rows, cols, colWidths, headerOffset, getCell, pendingCells, measureRowHeight, viewRange.firstRow, viewRange.lastRow, controlledRowHeights]);

  /** 关掉换行：清掉「自动算出来的折行高度」，但**保留用户手动拖出来的行高** */
  useEffect(() => {
    if (wrapEnabled) return;
    let changed = false;
    for (const row of [...heightsRef.current.keys()]) {
      if (manualRowsRef.current.has(row)) continue;
      heightsRef.current.delete(row);
      changed = true;
    }
    if (changed) setHeightsVersion((value) => value + 1);
  }, [wrapEnabled]);

  /**
   * 行高变化后的锚点保持：换行高度是「边滚边算」的，已加载行一旦变高，
   * 它下面所有行的位置都会平移 —— 不修正 scrollTop 的话，用户会看到内容整体跳动。
   * 这里以「视口顶部那一行」为锚：它的顶部 y 平移了多少，scrollTop 就补偿多少。
   */
  const prevRowIndexRef = useRef(rowIndex);
  useLayoutEffect(() => {
    const el = scrollerRef.current;
    const prev = prevRowIndexRef.current;
    prevRowIndexRef.current = rowIndex;
    if (!el || prev === rowIndex) return;
    const anchorRow = prev.virtualAt(el.scrollTop) + headerOffset;
    const delta = rowIndex.topOf(anchorRow) - prev.topOf(anchorRow);
    if (delta === 0) return;
    // 标记这次是程序化滚动：handleScroll 不会把它当成「用户滚动」去提交正在编辑的内容
    programmaticScrollRef.current = true;
    el.scrollTop = Math.max(0, el.scrollTop + delta);
    window.requestAnimationFrame(() => {
      programmaticScrollRef.current = false;
    });
  }, [rowIndex, headerOffset]);

  /**
   * 自动调整列宽（Excel 的自动调整列宽）：autoFitToken 变化时按**已加载**内容采样一次。
   * 取舍见 computeAutoFitWidths；采样读的是 readCellText（**待提交的编辑优先**）：刚敲进去还没保存的长串也算数。
   * 手动拖过宽度的列在 colWidths 合并时优先，因此自动调整不会覆盖它。
   *
   * **令牌只在真的采到内容时才消费**：CSV 的表格是 debounce 解析的、xlsx 的行窗口是异步加载的，
   * 工具栏点「自动调整列宽」时数据常常还没到 —— 那时若把令牌吃掉就再也不会补做，
   * 用户看到的就是「点了没反应」（用户实测的 CSV 自动调整列宽无效）。这里把 getCell / pendingCells
   * 放进依赖：数据到达（引用变化）后 effect 会再跑一次，令牌还没消费就补做一次。
   */
  const lastFitTokenRef = useRef<number | null>(null);
  /** 首次看到的令牌值 = 挂载时的状态：它那次属于「打开文件时自动拟合」，不算用户显式触发 */
  const initialFitTokenRef = useRef<number | null>(null);
  useEffect(() => {
    if (autoFitToken === undefined || empty) return;
    if (lastFitTokenRef.current === autoFitToken) return;
    if (initialFitTokenRef.current === null) initialFitTokenRef.current = autoFitToken;
    let sawContent = false;
    const widths = computeAutoFitWidths(cols, Math.min(rows, AUTOFIT_SAMPLE_ROWS), visibleRowList(viewRange, headerOffset), (row, col) => {
      const text = readCellText(callbacksRef.current, row, col);
      if (text.length > 0) sawContent = true;
      return text;
    });
    if (!sawContent) return; // 数据还没到：不消费令牌，等 getCell 换引用后补做
    lastFitTokenRef.current = autoFitToken;
    setFitWidths(widths);
    setLayoutVersion((value) => value + 1);
    /**
     * 用户**显式**触发（令牌变化）时，把变化报给父组件。
     * 为什么必须回调：受控列宽（用户拖过 / 撤销恢复回来的那一份存在父组件里）优先级高于自动调整结果，
     * 网格自己改不动它 —— 不回调的话「拖过一次之后，自动调整列宽就永远无效」（用户实测的 CSV 问题）。
     * 挂载时那次（initialFitTokenRef 记下的首个令牌值）不回调：否则一打开文件就凭空多出撤销步。
     */
    if (autoFitToken !== initialFitTokenRef.current) {
      const changes: Array<{ index: number; width: number }> = [];
      for (let col = 0; col < cols; col += 1) {
        const next = widths[col];
        // 拟合结果与当前生效宽度一致就不用报（避免无意义的撤销步）
        if (next === undefined || next === colWidths[col]) continue;
        changes.push({ index: col, width: next });
      }
      if (changes.length > 0) {
        const batched = callbacksRef.current.onColumnsResize;
        if (batched) batched(changes);
        else for (const change of changes) callbacksRef.current.onColumnResize?.(change.index, change.width);
      }
    }
    // 只认令牌：窗口/列数通过当前渲染读到的值取，不放进依赖（否则滚动一次就重算一遍）；
    // 但数据身份（getCell / pendingCells）必须放进依赖，否则「数据后到」永远不会补做
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoFitToken, empty, getCell, pendingCells]);

  /* ------------------ 手动调整列宽 / 行高（受控优先，未受控时留在会话内） ------------------ */

  /**
   * 「用户把某一列宽度改成多少」的落地。
   * - **回调先发**：只要父组件接了 onColumnResize 就一定收到（它负责记撤销栈）。这与是否受控无关 ——
   *   父组件第一次调整时往往还没传 columnWidths，那一手同样要能被撤销；
   * - 受控（父组件传了 columnWidths）：内部不写状态，等父组件把新值传回来才生效 ——
   *   这样网格内部不会留一份与 props 打架的状态；
   * - 未受控：写回内部表（会话内有效，与之前的行为完全一致）。
   */
  const applyColumnWidth = useCallback((col: number, width: number) => {
    callbacksRef.current.onColumnResize?.(col, width);
    if (callbacksRef.current.controlledColumns) return;
    manualColsRef.current.set(col, width);
    setColsVersion((value) => value + 1);
  }, []);

  /**
   * 「用户把某一行行高改成多少」的落地：受控语义同 applyColumnWidth。
   * 未受控时把该行记进 manualRowsRef（换行扫描不再改它，双击边界才恢复自动）。
   */
  const applyRowHeight = useCallback((row: number, height: number) => {
    callbacksRef.current.onRowResize?.(row, height);
    if (callbacksRef.current.controlledRows) return;
    heightsRef.current.set(row, height);
    manualRowsRef.current.add(row);
    setHeightsVersion((value) => value + 1);
  }, []);

  /**
   * 开始一次宽高调整。命中区在列标头右边界 / 行号栏下边界（见渲染里的 role="separator"），
   * 拖动期间**不触发选区、不进入编辑、不弹右键菜单** —— 那几个 handler 都在别的元素上。
   * 起始尺寸从 geomRef 现取（列宽/行高会随换行与手动调整变化），因此本回调引用稳定。
   */
  const beginResize = useCallback((kind: "col" | "row", index: number, clientX: number, clientY: number) => {
    const geom = geomRef.current;
    const startSize = kind === "col" ? geom.colWidths[index] ?? DEFAULT_COL_W : geom.rowIndex.heightOf(index);
    const rect = scrollerRef.current?.getBoundingClientRect();
    resizeRef.current = {
      kind,
      index,
      startClient: kind === "col" ? clientX : clientY,
      startSize,
      size: startSize,
      min: kind === "col" ? MIN_MANUAL_COL_W : MIN_MANUAL_ROW_H,
      max: kind === "col" ? MAX_MANUAL_COL_W : MAX_MANUAL_ROW_H,
      originLeft: rect?.left ?? 0,
      originTop: rect?.top ?? 0,
    };
    setResizeGuide({ kind, client: guideOffset(scrollerRef.current, kind, clientX, clientY) });
  }, []);

  /** 行号栏下边界起手（SheetRow 只需要 (row, clientX, clientY)） */
  const beginResizeRow = useCallback(
    (row: number, clientX: number, clientY: number) => beginResize("row", row, clientX, clientY),
    [beginResize],
  );

  /**
   * 一批行高变化的落地（一键自动调整行高用）。
   * - **一次调用只回调一次**：有 onRowsResize 就整批给它（父组件用一次 pushLayoutStep 记成一个撤销步），
   *   没传才退回逐行 onRowResize；
   * - 受控（父组件传了 rowHeights）时不写内部状态，等父组件把值传回来；
   * - 未受控时写内部表，并把这些行从 manualRowsRef 里移除 —— 一键调整的结果应当继续跟随内容
   *   （列宽变化、数据到达时由换行扫描重算），而不是像手动拖动那样被永久钉住。
   */
  const applyRowsHeight = useCallback((changes: Array<{ index: number; height: number }>) => {
    if (changes.length === 0) return;
    const live = callbacksRef.current;
    if (live.onRowsResize) live.onRowsResize(changes);
    else for (const change of changes) live.onRowResize?.(change.index, change.height);
    if (live.controlledRows) return;
    for (const change of changes) {
      if (change.height === ROW_H) heightsRef.current.delete(change.index);
      else heightsRef.current.set(change.index, change.height);
      manualRowsRef.current.delete(change.index);
    }
    setHeightsVersion((value) => value + 1);
  }, []);

  /**
   * 双击列边界：该列按内容自动调整。采样与 autoFitToken **完全一致**（前 200 行 + 当前可视窗口，
   * 含待提交编辑）—— 否则「滚到第 5000 行，屏幕上明明有内容，双击边界却算不出任何内容」，
   * 列宽会被设成默认值/更窄值，看起来像双击没反应。
   */
  const autoFitColumn = useCallback(
    (col: number) => {
      const widths = computeAutoFitWidths(
        cols,
        Math.min(rows, AUTOFIT_SAMPLE_ROWS),
        visibleRowList(viewRangeRef.current, geomRef.current.headerOffset),
        (row, index) => readCellText(callbacksRef.current, row, index),
      );
      const next = Math.max(MIN_MANUAL_COL_W, Math.min(MAX_MANUAL_COL_W, widths[col] ?? DEFAULT_COL_W));
      applyColumnWidth(col, next);
      setLayoutVersion((value) => value + 1);
    },
    [cols, rows, applyColumnWidth],
  );

  /** 双击行边界：该行按内容自动调整行高，并恢复「自动」（后续换行扫描可以再改它） */
  const autoFitRow = useCallback(
    (row: number) => {
      const measured = Math.max(MIN_MANUAL_ROW_H, Math.min(MAX_MANUAL_ROW_H, measureRowHeight(row)));
      // 与拖动同一条链路：先回调（父组件记撤销栈），受控时等 props 回传
      callbacksRef.current.onRowResize?.(row, measured);
      if (callbacksRef.current.controlledRows) {
        setLayoutVersion((value) => value + 1);
        return;
      }
      manualRowsRef.current.delete(row);
      if (measured === heightsRef.current.get(row)) return;
      if (measured === ROW_H) heightsRef.current.delete(row);
      else heightsRef.current.set(row, measured);
      setHeightsVersion((value) => value + 1);
      setLayoutVersion((value) => value + 1);
    },
    [measureRowHeight],
  );

  /**
   * 一键自动调整行高（Excel 的「自动调整行高」）：autoFitRowsToken 变化时，
   * 把**扫描窗口内当前已加载**的行按内容重算一次行高（与换行扫描同一个窗口，见 heightScanWindow）。
   *
   * 语义与双击单行边界一致（共用 measureRowHeight），但有三点是一键调整特有的：
   * 1) 只作用于**当时有内容可测**的行（已加载的单元格，或待提交编辑）：文件里本来没有这一格、
   *    用户刚输入还没保存的长串也算 —— 否则「输入长内容后一键自动调整行高没反应」；
   *    既没有单元格也没有待提交值的行保持默认，也不会把这次动作记在它们身上；
   * 2) 关闭 wrapText 时行高恒为默认（单行）—— 沿用双击边界的语义，不会把行撑高；
   * 3) 所有变化收集成**一批**交给 applyRowsHeight：父组件一次 pushLayoutStep 记成一个撤销步。
   * 令牌只消费一次；窗口里一行都没加载时不消费（等数据到达、getCell 换引用后补做一次）。
   * 与列宽的 autoFitToken 各用各的 ref，互不影响。
   */
  const lastFitRowsTokenRef = useRef<number | null>(null);
  useEffect(() => {
    if (autoFitRowsToken === undefined || empty) return;
    if (lastFitRowsTokenRef.current === autoFitRowsToken) return;
    const scanWindow = heightScanWindow(viewRange, headerOffset, rows);
    const changes: Array<{ index: number; height: number }> = [];
    let sawLoaded = false;
    for (let row = scanWindow.first; row <= scanWindow.last; row += 1) {
      // 与换行扫描共用同一个测量入口：**待提交编辑也算有内容**（文件里本来没有这一格、
      // 用户刚输入还没保存的长串同样要能被一键调整），没有可测内容的行才跳过。
      const measured = measureRow(row);
      if (!measured.hasContent) continue;
      sawLoaded = true;
      const next = measured.height;
      const prev = effectiveRowHeights.get(row) ?? ROW_H;
      if (next === prev) continue;
      changes.push({ index: row, height: next });
    }
    if (!sawLoaded) return; // 一行都没加载：不消费令牌，数据到达后（getCell 换引用）再补做
    lastFitRowsTokenRef.current = autoFitRowsToken;
    if (changes.length === 0) return;
    applyRowsHeight(changes);
    setLayoutVersion((value) => value + 1);
    // 只认令牌：窗口 / 列宽 / 行高表都按当前渲染读到的值取，不放进依赖（否则滚动一次就重算一遍）
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoFitRowsToken, empty, getCell, pendingCells, measureRow, applyRowsHeight]);

  /**
   * 调整会话的 window 监听：拖动时实时改宽度/高度（受控模式下即实时回调父组件），
   * 松手才通知父组件 onLayoutChange（拖动过程中每一帧都回调会把统计/滚动条刷爆）。
   * 与选区拖拽是两套独立会话（resizeRef vs dragRef），互不干扰。
   */
  useEffect(() => {
    const onMove = (event: MouseEvent) => {
      const session = resizeRef.current;
      if (!session) return;
      const client = session.kind === "col" ? event.clientX : event.clientY;
      const next = Math.max(session.min, Math.min(session.max, Math.round(session.startSize + (client - session.startClient))));
      setResizeGuide({
        kind: session.kind,
        client: session.kind === "col" ? client - session.originLeft : client - session.originTop,
      });
      if (next === session.size) return;
      session.size = next;
      // applyColumnWidth / applyRowHeight 是稳定引用（useCallback([])），这里可以直接用
      if (session.kind === "col") applyColumnWidth(session.index, next);
      else applyRowHeight(session.index, next);
    };
    const onUp = () => {
      if (!resizeRef.current) return;
      resizeRef.current = null;
      setResizeGuide(null);
      // 尺寸定了才通知父组件：拖动过程中每一帧都回调会把统计/滚动条刷爆
      setLayoutVersion((value) => value + 1);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    window.addEventListener("blur", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      window.removeEventListener("blur", onUp);
    };
  }, []);

  /** 列宽/行高真的变了 → 通知父组件（内容总尺寸可能变了，滚动条、统计要跟着刷新） */
  useEffect(() => {
    if (layoutVersion === 0) return; // 首次挂载不算「变化」
    callbacksRef.current.onLayoutChange?.();
  }, [layoutVersion]);

  /**
   * 公式引用框：父组件把「当前单元格公式引用到的区域」传进来，网格只负责画框。
   * 只画落在**当前渲染窗口**内的部分（不因此多渲染任何行、任何列），
   * 且只有「真正的边」才描线 —— 被窗口裁掉的那一侧不描，免得看起来引用就在那里结束。
   * 多个引用按顺序循环取色（Excel 里不同引用用不同颜色区分）。
   */
  const formulaFrames = useMemo(() => {
    if (!formulaRanges || formulaRanges.length === 0 || empty) return [];
    const topRow = viewRange.firstRow + headerOffset;
    const bottomRow = viewRange.lastRow + headerOffset;
    const frames: Array<{
      key: string;
      left: number;
      top: number;
      width: number;
      height: number;
      showTop: boolean;
      showRight: boolean;
      showBottom: boolean;
      showLeft: boolean;
      tone: string;
    }> = [];
    formulaRanges.forEach((range, index) => {
      // 端点可能是反向的（父组件把拖选出来的区域原样传进来）：几何与「哪条边真的在窗口里」
      // 都按归一化后的值算 —— 否则倒序区域会在被窗口裁掉的那一侧多描一条边，
      // 看起来引用正好在那里结束（正序区域不会）。
      const row0 = Math.min(range.startRow, range.endRow);
      const row1 = Math.max(range.startRow, range.endRow);
      const col0 = Math.min(range.startCol, range.endCol);
      const col1 = Math.max(range.startCol, range.endCol);
      const startRow = Math.max(row0, topRow);
      const endRow = Math.min(row1, bottomRow);
      const startCol = Math.max(col0, viewRange.firstCol);
      const endCol = Math.min(col1, viewRange.lastCol);
      if (startRow > endRow || startCol > endCol) return;
      const left = colOffsets[startCol];
      const right = colOffsets[endCol] + colWidths[endCol];
      const top = rowIndex.topOf(startRow);
      const bottom = rowIndex.topOf(endRow) + rowIndex.heightOf(endRow);
      frames.push({
        key: `${index}:${startRow}:${endRow}:${startCol}:${endCol}`,
        left,
        top,
        width: Math.max(1, right - left),
        height: Math.max(1, bottom - top),
        showTop: row0 >= topRow,
        showRight: col1 <= viewRange.lastCol,
        showBottom: row1 <= bottomRow,
        showLeft: col0 >= viewRange.firstCol,
        tone: FORMULA_TONES[index % FORMULA_TONES.length],
      });
    });
    return frames;
  }, [
    formulaRanges,
    empty,
    headerOffset,
    rowIndex,
    colOffsets,
    colWidths,
    viewRange.firstRow,
    viewRange.lastRow,
    viewRange.firstCol,
    viewRange.lastCol,
  ]);

  // 受控 activeCell / selection 的同步 effect 已挪到 ensureVisible 之后
  // （程序化定位要真的滚过去，就得能在 effect 里直接引用 ensureVisible）。

  /**
   * 把单元格滚入可视区（纵向避开顶部冻结区，横向避开左侧行号栏 + 冻结列）。
   * 目标格本身落在**冻结区**里时它一直可见，不滚动（否则用户点一下冻结行/列，整个表格会莫名跳一下）。
   */
  const ensureVisible = useCallback(
    (row: number, col: number) => {
      const el = scrollerRef.current;
      if (!el) return;
      if (row >= headerOffset) {
        // 行高索引给出的「这一行的顶部 y 与高度」：开启换行后行高可能远大于 ROW_H
        const y = stickyH + rowIndex.topOf(row);
        const rowH = rowIndex.heightOf(row);
        if (y < el.scrollTop + stickyH) el.scrollTop = Math.max(0, y - stickyH);
        else if (y + rowH > el.scrollTop + el.clientHeight) el.scrollTop = y + rowH - el.clientHeight;
      }

      if (col >= freezeColCount) {
        const x = colOffsets[col] ?? GUTTER_W;
        const width = colWidths[col] ?? DEFAULT_COL_W;
        if (x < el.scrollLeft + frozenWidth) el.scrollLeft = Math.max(0, x - frozenWidth);
        else if (x + width > el.scrollLeft + el.clientWidth) el.scrollLeft = x + width - el.clientWidth;
      }
    },
    [stickyH, rowIndex, colOffsets, colWidths, headerOffset, freezeColCount, frozenWidth],
  );

  // 受控同步：undefined 表示非受控（网格自己管选中），null 表示显式清空；
  // 值相同时复用原对象，避免父组件每次传新字面量导致多余渲染。
  // 选区顺带跟着走：只有当选区还「就是活动格本身」（1×1 或还没有选区）时才跟随，
  // 多格选区不会被外部改 activeCell 打散（要清空请显式传 selection={null}）。
  // 放在 ensureVisible 之后：程序化定位（查找跳转等）要真的滚过去。
  useEffect(() => {
    if (activeCell === undefined) return;
    if (activeCell === null) {
      setActive(null); // 已是 null 时 React 会跳过这次渲染
      if (!callbacksRef.current.selectionControlled) {
        anchorRef.current = null;
        setRange(null);
      }
      return;
    }
    if (rows === 0 || cols === 0) return;
    const row = Math.max(0, Math.min(rows - 1, activeCell.row));
    const col = Math.max(0, Math.min(cols - 1, activeCell.col));
    const prev = activeRef.current;
    setActive((old) => (old && old.row === row && old.col === col ? old : { row, col }));
    // 只有「确实换了格」才滚动：父组件反复传同一个值时不能跟用户的手动滚动打架
    if (!prev || prev.row !== row || prev.col !== col) ensureVisible(row, col);
    if (callbacksRef.current.selectionControlled || dragRef.current) return;
    const prevRange = rangeRef.current;
    const trivial = prevRange === null || (prev !== null && sameRange(prevRange, singleRange(prev.row, prev.col)));
    if (!trivial) return;
    anchorRef.current = { row, col };
    const next = singleRange(row, col);
    setRange((old) => (sameRange(old, next) ? old : next));
  }, [activeCell, ensureVisible, rows, cols]);

  // 受控选区：父组件传具体值或 null 时以它为准（拖拽进行中不打断自己；值等价时不重渲染）
  useEffect(() => {
    if (selection === undefined) return;
    if (dragRef.current) return;
    if (selection === null) {
      setRange(null);
      return;
    }
    if (rows === 0 || cols === 0) return;
    const next = clampRange(selection, rows, cols);
    if (!next) return;
    anchorRef.current = { row: next.startRow, col: next.startCol };
    setRange((prev) => (sameRange(prev, next) ? prev : next));
  }, [selection, rows, cols]);

  /**
   * 查找跳转的滚动请求：**只滚，不动任何选中状态**。
   * 用户报告「点查找之后选区高亮就没了」—— 根因是跳转时借用了「设置活动单元格」来滚动（会移动焦点框）；
   * 这里走独立的 token：token 变化才执行一次，选区 / 活动格 / 焦点框一概不碰。
   * 目标落在冻结区里时不滚动（ensureVisible 的既有语义）。滚动标记为程序化：不会顺手提交正在编辑的内容。
   */
  const lastScrollTokenRef = useRef<number | null>(null);
  useEffect(() => {
    if (!scrollTarget) return;
    if (lastScrollTokenRef.current === scrollTarget.token) return;
    lastScrollTokenRef.current = scrollTarget.token;
    const el = scrollerRef.current;
    if (!el) return;
    programmaticScrollRef.current = true;
    ensureVisible(scrollTarget.row, scrollTarget.col);
    window.requestAnimationFrame(() => {
      programmaticScrollRef.current = false;
    });
  }, [scrollTarget, ensureVisible]);


  /* ------------------------------ 选区：状态机 ------------------------------ */

  /**
   * 统一的选区出口。语义：
   * - 值没变 → 不重渲染、不回调 onSelectionChange（但 commit 是「手势结束」信号，照样回调）；
   * - 值变了 → 更新内部状态 + 回调 onSelectionChange（commit 时再补一次 onSelectionCommit）。
   * rangeRef 在这里同步跟手，保证同一轮事件里连续几次判断都基于最新值。
   */
  const emitSelection = useCallback((next: SheetRange | null, commit: boolean) => {
    if (!sameRange(rangeRef.current, next)) {
      rangeRef.current = next;
      setRange(next);
      callbacksRef.current.onSelectionChange?.(next);
    }
    if (commit) callbacksRef.current.onSelectionCommit?.(next);
  }, []);

  /** 点选 / 键盘移动：锚点跟着走，选区收缩成 1×1（不是 null，父组件用一个 range 就能表达统计口径） */
  const selectSingle = useCallback(
    (row: number, col: number, commit: boolean) => {
      anchorRef.current = { row, col };
      emitSelection(singleRange(row, col), commit);
    },
    [emitSelection],
  );

  /** Shift 扩展：从锚点拉到目标格；没有锚点时退化成点选 */
  const extendSelection = useCallback(
    (row: number, col: number, commit: boolean) => {
      const anchor = anchorRef.current;
      if (!anchor) {
        selectSingle(row, col, commit);
        return;
      }
      emitSelection(rangeOf(anchor, { row, col }), commit);
    },
    [emitSelection, selectSingle],
  );

  /* ------------------------------ 公式拾取（Excel 的点选引用） ------------------------------ */

  /** 编辑框注册/注销改写能力（拾取结束要把引用写进编辑框） */
  const bindEditorApi = useCallback((api: CellEditorApi | null) => {
    editorApiRef.current = api;
  }, []);

  /** 编辑框文本是否以 = 开头：只在真的翻转时 setState，避免每次击键都重渲染网格 */
  const notifyFormulaMode = useCallback((active: boolean) => {
    if (formulaModeRef.current === active) return;
    formulaModeRef.current = active;
    if (!active) {
      setPicking(false);
      pickStartRef.current = null;
    }
  }, []);

  /** 读「是否正在拾取」：Esc 的最高优先级（取消拾取而不是退出编辑） */
  const readPicking = useCallback(() => dragRef.current?.pick !== undefined, []);

  /**
   * 开始一次拾取：按下鼠标时**不改活动格、不触发编辑、不移动焦点**，
   * 只把拖出来的选区当作「引用」；松开鼠标时由 onUp 交给编辑框或父组件。
   */
  const beginPick = useCallback(
    (row: number, col: number, mode: "cell" | "row" | "col", target: "editor" | "external", clientX: number, clientY: number) => {
      if (empty) return;
      pickStartRef.current = rangeRef.current;
      const single: SheetRange =
        mode === "col"
          ? { startRow: 0, startCol: col, endRow: rows - 1, endCol: col }
          : mode === "row"
            ? { startRow: row, startCol: 0, endRow: row, endCol: cols - 1 }
            : singleRange(row, col);
      dragRef.current = { anchor: { row, col }, mode, pick: target, clientX, clientY, range: single };
      setPicking(true);
      emitSelection(single, false);
    },
    [cols, emitSelection, empty, rows],
  );

  /** Esc 取消拾取：恢复拾取前的选区，编辑框保持焦点与内容不变 */
  const cancelPick = useCallback(() => {
    const drag = dragRef.current;
    if (!drag?.pick) return;
    dragRef.current = null;
    setPicking(false);
    const back = pickStartRef.current;
    pickStartRef.current = null;
    if (back) emitSelection(back, false);
  }, [emitSelection]);

  /** 整列：全部行 × [colA, colB]（列标头点击 / 拖过列标头） */
  const selectColumns = useCallback(    (colA: number, colB: number, commit: boolean) => {
      if (empty) return;
      emitSelection(
        { startRow: 0, startCol: Math.min(colA, colB), endRow: rows - 1, endCol: Math.max(colA, colB) },
        commit,
      );
    },
    [emitSelection, empty, rows],
  );

  /** 整行：全部列 × [rowA, rowB]（行号栏点击 / 拖过行号栏） */
  const selectRows = useCallback(
    (rowA: number, rowB: number, commit: boolean) => {
      if (empty) return;
      emitSelection(
        { startRow: Math.min(rowA, rowB), startCol: 0, endRow: Math.max(rowA, rowB), endCol: cols - 1 },
        commit,
      );
    },
    [cols, emitSelection, empty],
  );

  /** 全选：Ctrl+A 与左上角（活动格不动，与 Excel 一致） */
  const selectAll = useCallback(
    (commit: boolean) => {
      if (empty) return;
      anchorRef.current = { row: 0, col: 0 };
      emitSelection({ startRow: 0, startCol: 0, endRow: rows - 1, endCol: cols - 1 }, commit);
    },
    [cols, emitSelection, empty, rows],
  );

  /**
   * 单元格按下：Shift 从锚点扩展，否则重开一个以该格为锚点的 1×1 选区；同时建立拖拽会话。
   * 拖拽期间活动格固定在锚点、不逐帧回调 onActiveCell，只有 onSelectionChange 会连续触发。
   */
  const handleCellMouseDown = useCallback(
    (row: number, col: number, shiftKey: boolean, clientX: number, clientY: number) => {
      if (empty) return;
      // 公式拾取模式：编辑框里的文本以 = 开头（或内容栏正在编辑公式）时，点/拖网格只用来「选引用」——
      // 不改活动格、不切编辑格、不抢焦点（单元格的 mousedown 已经 preventDefault）
      if (editingRef.current && formulaModeRef.current) {
        beginPick(row, col, "cell", "editor", clientX, clientY);
        return;
      }
      if (externalPick) {
        beginPick(row, col, "cell", "external", clientX, clientY);
        return;
      }
      // 先要焦点：正在进行的编辑会在这里被 blur 提交掉，之后再改选区才不会打架
      scrollerRef.current?.focus({ preventScroll: true });
      if (shiftKey) {
        // Excel：Shift+点击既扩展选区，也把活动格（焦点端）挪到点击处
        extendSelection(row, col, false);
        setActive((prev) => (prev && prev.row === row && prev.col === col ? prev : { row, col }));
        callbacksRef.current.onActiveCell?.(row, col);
      } else {
        setActive({ row, col });
        callbacksRef.current.onActiveCell?.(row, col);
        selectSingle(row, col, false);
      }
      ensureVisible(row, col);
      const anchor = shiftKey ? anchorRef.current ?? { row, col } : { row, col };
      dragRef.current = { anchor, mode: "cell", clientX, clientY, range: rangeOf(anchor, { row, col }) };
    },
    [beginPick, empty, ensureVisible, extendSelection, externalPick, selectSingle],
  );

  /**
   * 列标头按下：整列（Shift 从锚点列扩展成多列）。
   * 活动格落在第一条可滚动行（有冻结表头时跳过它），锚点用第 0 行，方便再 Shift 扩展。
   */
  const handleColumnMouseDown = useCallback(
    (col: number, shiftKey: boolean, clientX: number, clientY: number) => {
      if (empty) return;
      // 公式拾取：列标头拖出的是整列引用（B:B）
      if (editingRef.current && formulaModeRef.current) {
        beginPick(Math.min(headerOffset, rows - 1), col, "col", "editor", clientX, clientY);
        return;
      }
      if (externalPick) {
        beginPick(Math.min(headerOffset, rows - 1), col, "col", "external", clientX, clientY);
        return;
      }
      scrollerRef.current?.focus({ preventScroll: true });
      const anchorCol = shiftKey ? anchorRef.current?.col ?? col : col;
      const row = Math.min(headerOffset, rows - 1);
      setActive((prev) => (prev && prev.row === row && prev.col === col ? prev : { row, col }));
      callbacksRef.current.onActiveCell?.(row, col);
      selectColumns(anchorCol, col, false);
      anchorRef.current = { row: 0, col: anchorCol };
      dragRef.current = {
        anchor: { row: 0, col: anchorCol },
        mode: "col",
        clientX,
        clientY,
        range: {
          startRow: 0,
          startCol: Math.min(anchorCol, col),
          endRow: rows - 1,
          endCol: Math.max(anchorCol, col),
        },
      };
    },
    [beginPick, empty, externalPick, headerOffset, rows, selectColumns],
  );

  /** 行号栏按下：整行（Shift 从锚点行扩展成多行） */
  const handleRowMouseDown = useCallback(
    (row: number, shiftKey: boolean, clientX: number, clientY: number) => {
      if (empty) return;
      // 公式拾取：行号栏拖出的是整行引用（3:3）
      if (editingRef.current && formulaModeRef.current) {
        beginPick(row, 0, "row", "editor", clientX, clientY);
        return;
      }
      if (externalPick) {
        beginPick(row, 0, "row", "external", clientX, clientY);
        return;
      }
      scrollerRef.current?.focus({ preventScroll: true });
      const anchorRow = shiftKey ? anchorRef.current?.row ?? row : row;
      setActive((prev) => (prev && prev.row === row && prev.col === 0 ? prev : { row, col: 0 }));
      callbacksRef.current.onActiveCell?.(row, 0);
      selectRows(anchorRow, row, false);
      anchorRef.current = { row: anchorRow, col: 0 };
      dragRef.current = {
        anchor: { row: anchorRow, col: 0 },
        mode: "row",
        clientX,
        clientY,
        range: {
          startRow: Math.min(anchorRow, row),
          startCol: 0,
          endRow: Math.max(anchorRow, row),
          endCol: cols - 1,
        },
      };
    },
    [beginPick, cols, empty, externalPick, selectRows],
  );

  /** 左上角：等价于 Ctrl+A（一次点击就完成 change + commit） */
  const handleCornerMouseDown = useCallback(
    (event: React.MouseEvent<HTMLDivElement>) => {
      // 拾取中：左上角不抢焦点、不改选区（否则内容栏的编辑框会失焦）
      if (readPicking()) {
        event.preventDefault();
        return;
      }
      scrollerRef.current?.focus({ preventScroll: true });
      selectAll(true);
    },
    [readPicking, selectAll],
  );

  /* ------------------------------ 右键菜单 ------------------------------ */

  /**
   * 右键公共逻辑：Excel 规则 —— 目标不在当前选区内时先把选区/活动格移过去，再回调父组件。
   * 返回值 = 父组件是否接了回调（由行/表头那边决定要不要 preventDefault 掉系统菜单：
   * 没接回调就保留浏览器默认菜单，免得右键彻底没反应）。
   * 网格自己不渲染任何菜单。
   */
  const handleCellContextMenu = useCallback(
    (row: number, col: number, clientX: number, clientY: number): boolean => {
      // 拾取过程中抑制右键菜单（返回 true 让调用方 preventDefault 掉系统菜单）
      if (readPicking()) return true;
      const handler = callbacksRef.current.onCellContextMenu;
      if (!handler || empty) return false;
      if (!inRange(rangeRef.current, row, col)) {
        scrollerRef.current?.focus({ preventScroll: true });
        setActive({ row, col });
        callbacksRef.current.onActiveCell?.(row, col);
        selectSingle(row, col, true);
        ensureVisible(row, col);
      }
      handler({ row, col, kind: "cell" }, { x: clientX, y: clientY });
      return true;
    },
    [empty, ensureVisible, selectSingle],
  );

  /** 行号栏右键：目标行没被整行选中就先选中整行（Excel 行为），col 用当前活动列 */
  const handleRowContextMenu = useCallback(
    (row: number, clientX: number, clientY: number): boolean => {
      // 拾取过程中抑制右键菜单
      if (readPicking()) return true;
      const handler = callbacksRef.current.onCellContextMenu;
      if (!handler || empty) return false;
      const current = rangeRef.current;
      const wholeRow = inRange(current, row, 0) && inRange(current, row, cols - 1);
      const col = Math.max(0, Math.min(cols - 1, activeRef.current?.col ?? current?.startCol ?? 0));
      if (!wholeRow) {
        setActive((prev) => (prev && prev.row === row && prev.col === col ? prev : { row, col }));
        callbacksRef.current.onActiveCell?.(row, col);
        selectRows(row, row, true);
        ensureVisible(row, col);
      }
      handler({ row, col, kind: "rowHeader" }, { x: clientX, y: clientY });
      return true;
    },
    [cols, empty, ensureVisible, selectRows],
  );

  /** 列标头右键：目标列没被整列选中就先选中整列，row 用当前活动行 */
  const handleColumnContextMenu = useCallback(
    (col: number, clientX: number, clientY: number): boolean => {
      // 拾取过程中抑制右键菜单
      if (readPicking()) return true;
      const handler = callbacksRef.current.onCellContextMenu;
      if (!handler || empty) return false;
      const current = rangeRef.current;
      const wholeCol = inRange(current, 0, col) && inRange(current, rows - 1, col);
      const row = Math.max(0, Math.min(rows - 1, activeRef.current?.row ?? current?.startRow ?? 0));
      if (!wholeCol) {
        setActive((prev) => (prev && prev.row === row && prev.col === col ? prev : { row, col }));
        callbacksRef.current.onActiveCell?.(row, col);
        selectColumns(col, col, true);
        ensureVisible(row, col);
      }
      handler({ row, col, kind: "colHeader" }, { x: clientX, y: clientY });
      return true;
    },
    [empty, ensureVisible, rows, selectColumns],
  );

  /**
   * client 坐标 → 数据行列（换算到内容坐标；指针落在冻结区里时按冻结区自己的坐标算）。
   * 选区拖拽与拖拽填充共用同一个换算，只有一处实现。
   */
  const pointToCell = useCallback(
    (clientX: number, clientY: number): { row: number; col: number } | null => {
      const el = scrollerRef.current;
      const geom = geomRef.current;
      if (!el || geom.rows === 0 || geom.cols === 0) return null;
      const rect = el.getBoundingClientRect();
      // 横向：指针在「行号栏 + 冻结列」这块区域里时，它对应的就是屏幕上那几列（不随 scrollLeft 走）
      const localX = clientX - rect.left;
      const x = localX < geom.frozenWidth ? localX : localX + el.scrollLeft;
      const col = clampIndex(colAtX(x, geom.colOffsets), geom.cols - 1);
      // 纵向：冻结行是 sticky 的，它们悬在正文上方，指针压在上面时按冻结区内部的 y 找到具体那一行；
      // 其余情况先把指针夹到正文区上边界（拖到网格上方 = 选到当前第一可见行），再按行高折算。
      const bodyTop = rect.top + geom.stickyH;
      if (clientY < bodyTop) {
        if (geom.headerOffset <= 0) {
          // 没有冻结行：夹到正文第一行
        } else {
          let offset = clientY - rect.top - HEADER_H;
          for (let row = 0; row < geom.headerOffset; row += 1) {
            const h = geom.rowIndex.heightOf(row);
            if (offset < h) return { row, col };
            offset -= h;
          }
          return { row: geom.headerOffset - 1, col };
        }
      }
      const y = Math.max(clientY, bodyTop) - rect.top + el.scrollTop;
      // 用行高索引换算：可高行的区域也能落在正确的行上（未开启换行时等价于原来的 y / ROW_H）
      const row = clampIndex(geom.rowIndex.virtualAt(y - geom.stickyH) + geom.headerOffset, geom.rows - 1);
      return { row, col };
    },
    [],
  );
  /** 供「只挂一次监听器」的会话读最新实现 */
  const pointToCellRef = useRef(pointToCell);
  pointToCellRef.current = pointToCell;

  /* ------------------ 拖拽自动填充（Excel 的填充柄） ------------------ */

  /**
   * 选区右下角的填充柄：只在**可编辑**且父组件接了 `onFillRange` 时显示
   * （只读网格不显示，避免误操作；不传回调时完全不渲染 → DOM 逐字节一致）。
   */
  const fillCorner = useMemo(() => {
    if (!onFillRange || !editable || empty || !sel) return null;
    return { row: sel.endRow, col: sel.endCol };
  }, [onFillRange, editable, empty, sel]);
  /** 拖拽中的虚线预览（null = 没在拖）；source 固定，target 随指针单向扩展 */
  const [fillPreview, setFillPreview] = useState<{ source: SheetRange; target: SheetRange } | null>(null);
  /** 会话里的权威状态：目标区域同步写在 ref 里，松手时不依赖「预览已经渲染过」 */
  const fillDragRef = useRef<{ source: SheetRange; target: SheetRange } | null>(null);

  /** 按下填充柄：建立会话（**不改选区、不改内容、不进入编辑**） */
  const beginFillDrag = useCallback((row: number, col: number) => {
    const source = rangeRef.current ?? singleRange(row, col);
    fillDragRef.current = { source, target: source };
    setFillPreview({ source, target: source });
  }, []);

  useEffect(() => {
    const onMove = (event: MouseEvent) => {
      const session = fillDragRef.current;
      if (!session) return;
      const point = pointToCellRef.current(event.clientX, event.clientY);
      if (!point) return;
      const target = extendFillTarget(session.source, point);
      session.target = target;
      // 目标没变就不 setState（拖动期间每一帧都渲染太浪费）
      setFillPreview((prev) => (prev && sameRange(prev.target, target) ? prev : { source: session.source, target }));
    };
    const onUp = () => {
      const session = fillDragRef.current;
      fillDragRef.current = null;
      setFillPreview(null);
      if (!session) return;
      // 拖回原选区（没有扩展）→ 不回调
      if (sameRange(session.source, session.target)) return;
      callbacksRef.current.onFillRange?.(session.source, session.target);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      // 拖拽期间按 Esc 取消：不回调
      if (event.key !== "Escape" || !fillDragRef.current) return;
      fillDragRef.current = null;
      setFillPreview(null);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    window.addEventListener("keydown", onKeyDown, true);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      window.removeEventListener("keydown", onKeyDown, true);
    };
  }, []);

  /**
   * 拖拽会话的 window 级监听：指针可能跑到网格之外（甚至划过表头），
   * 所以 mousemove / mouseup 挂在 window 上且**只挂一次**，内部靠 dragRef / geomRef 读最新状态，
   * 不随渲染重建监听器。指针贴到视口边缘时按帧自动滚动；滚动量不再变化（已经到边）就停，避免空转。
   */
  useEffect(() => {
    /** 指针位置 → 新选区；还在同一格时不 setState、不回调（拖拽期间的主要开销都在这里省掉） */
    const applyPointer = (clientX: number, clientY: number) => {
      const drag = dragRef.current;
      const point = drag ? pointToCellRef.current(clientX, clientY) : null;
      if (!drag || !point) return;
      drag.clientX = clientX;
      drag.clientY = clientY;
      const geom = geomRef.current;
      const next: SheetRange =
        drag.mode === "col"
          ? {
              startRow: 0,
              startCol: Math.min(drag.anchor.col, point.col),
              endRow: geom.rows - 1,
              endCol: Math.max(drag.anchor.col, point.col),
            }
          : drag.mode === "row"
            ? {
                startRow: Math.min(drag.anchor.row, point.row),
                startCol: 0,
                endRow: Math.max(drag.anchor.row, point.row),
                endCol: geom.cols - 1,
              }
            : rangeOf(drag.anchor, point);
      if (sameRange(drag.range, next)) return;
      drag.range = next;
      emitSelection(next, false); // emitSelection 是稳定引用（useCallback([])），这里可以直接用
    };

    let raf: number | null = null;
    const EDGE = 24;
    const schedule = () => {
      if (raf === null) raf = window.requestAnimationFrame(tick);
    };
    const tick = () => {
      raf = null;
      const drag = dragRef.current;
      const el = scrollerRef.current;
      if (!drag || !el) return;
      const rect = el.getBoundingClientRect();
      if (rect.bottom - rect.top <= 0) return; // 没有真实布局（首帧 / 无头环境）时不自动滚动
      let dy = 0;
      let dx = 0;
      // 垂直自动滚动步长取「当前第一可见行」的高度：换行后一行可能远高于 ROW_H，
      // 固定 24px 会让人以为拖到边缘几乎不动（未开启换行时它就是 ROW_H，行为与以前一致）
      const geom = geomRef.current;
      const stepH = Math.max(ROW_H, geom.rowIndex.heightOf(geom.rowIndex.virtualAt(el.scrollTop) + geom.headerOffset));
      if (drag.clientY < rect.top + geom.stickyH + EDGE) dy = -stepH;
      else if (drag.clientY > rect.bottom - EDGE) dy = stepH;
      if (drag.clientX < rect.left + GUTTER_W + EDGE) dx = -AUTO_SCROLL_STEP_X;
      else if (drag.clientX > rect.right - EDGE) dx = AUTO_SCROLL_STEP_X;
      if (dx === 0 && dy === 0) return;
      const beforeTop = el.scrollTop;
      const beforeLeft = el.scrollLeft;
      el.scrollTop = beforeTop + dy;
      el.scrollLeft = beforeLeft + dx;
      if (el.scrollTop === beforeTop && el.scrollLeft === beforeLeft) return; // 已经到边
      applyPointer(drag.clientX, drag.clientY); // 滚动后指针下的格子变了
      schedule();
    };

    const onMove = (event: MouseEvent) => {
      if (!dragRef.current) return;
      applyPointer(event.clientX, event.clientY);
      schedule(); // 指针贴边时由它起自动滚动循环；离开边缘后 tick 自己返回
    };
    const onUp = () => {
      const drag = dragRef.current;
      if (!drag) return;
      dragRef.current = null;
      if (drag.pick === "editor") {
        // 单元格内编辑框的拾取：把选区写成引用（B3 / B3:B4 / B:B / 3:3）交给编辑框插入或替换
        setPicking(false);
        pickStartRef.current = null;
        editorApiRef.current?.applyReference(rangeReference(drag.range, drag.mode));
        return; // 拾取不算一次「选区提交」，不回调 onSelectionCommit
      }
      if (drag.pick === "external") {
        // 内容栏的拾取：引用文本由父组件自己转换，这里只回 0 起坐标
        setPicking(false);
        pickStartRef.current = null;
        callbacksRef.current.onPickReference?.(drag.range);
        return;
      }
      callbacksRef.current.onSelectionCommit?.(drag.range); // 一次手势只提交一次
    };

    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    window.addEventListener("blur", onUp); // 指针在窗口外松开时收尾，避免拖拽状态卡住
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      window.removeEventListener("blur", onUp);
      if (raf !== null) window.cancelAnimationFrame(raf);
    };
  }, []);

  /** 键盘移动活动格：extend（Shift）保持锚点并扩展选区，否则选区收缩成新的 1×1 并提交一次 */
  const moveActive = useCallback(
    (row: number, col: number, extend = false) => {
      const nextRow = Math.max(0, Math.min(rows - 1, row));
      const nextCol = Math.max(0, Math.min(cols - 1, col));
      setActive({ row: nextRow, col: nextCol });
      callbacksRef.current.onActiveCell?.(nextRow, nextCol);
      if (extend) extendSelection(nextRow, nextCol, true);
      else selectSingle(nextRow, nextCol, true);
      ensureVisible(nextRow, nextCol);
    },
    [cols, ensureVisible, extendSelection, rows, selectSingle],
  );

  /* ------------------------------ 编辑：进入/提交/取消 ------------------------------ */

  /** 某格当前的显示文本：待提交的编辑优先，其次原始数据，都没有则是空串 */
  const displayText = useCallback((row: number, col: number): string => {
    const live = callbacksRef.current;
    return live.pendingCells?.get(`${row},${col}`) ?? live.getCell(row, col)?.v ?? "";
  }, []);

  /**
   * 父组件没接 onCopySelection 时的兜底：把选区序列化成 TSV 写剪贴板（单格时就是该格文本）。
   * 内容按「屏幕上看到的」来（待提交编辑优先）。规模保护：超过 2 万格直接放弃，
   * 否则 Ctrl+A 复制百万行表会把内存和剪贴板打爆。
   */
  const copyRangeFallback = useCallback(
    (range: SheetRange) => {
      if (rangeCellCount(range) > 20000) return;
      const lines: string[] = [];
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        const cells: string[] = [];
        for (let col = range.startCol; col <= range.endCol; col += 1) {
          const text = displayText(row, col);
          cells.push(/[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);
        }
        lines.push(cells.join("\t"));
      }
      try {
        void navigator.clipboard?.writeText(lines.join("\n")).catch(() => {});
      } catch {
        /* 剪贴板不可用（非安全上下文等）时静默忽略 */
      }
    },
    [displayText],
  );

  const bindEditorInput = useCallback((el: HTMLInputElement | null) => {
    editorInputRef.current = el;
  }, []);

  /**
   * 进入编辑。seed 供「选中后直接敲字符」使用：该字符成为初始内容（替换原值）。
   * 只读 ref，引用稳定，双击/F2/Enter 三个入口共用。
   */
  const beginEditing = useCallback(
    (row: number, col: number, caret: EditCaret, seed?: string) => {
      if (!callbacksRef.current.editable) return; // 只读：永不进入编辑态
      if (rows === 0 || cols === 0) return;
      const r = Math.max(0, Math.min(rows - 1, row));
      const c = Math.max(0, Math.min(cols - 1, col));
      const prev = activeRef.current;
      if (!prev || prev.row !== r || prev.col !== c) {
        // 编辑的那一格同时成为选中格（正常情况下双击前 mousedown 已经选中，这里是键盘入口的兜底），
        // 与 setEditing 落在同一批更新里，只多渲染一次
        setActive({ row: r, col: c });
        callbacksRef.current.onActiveCell?.(r, c);
      }
      const next: EditingState = { row: r, col: c, text: seed ?? displayText(r, c), caret };
      editingRef.current = next; // 同步写入：紧接着的同一轮事件里就能看见编辑态
      setEditing(next);
    },
    [cols, displayText, rows],
  );

  /** 双击进入编辑：初始内容全选，直接敲字即可替换 */
  const startEditFromCell = useCallback(
    (row: number, col: number) => {
      // 公式拾取模式：双击也只是把引用选进来，不要把编辑切到别的格去
      if (editingRef.current && formulaModeRef.current) return;
      beginEditing(row, col, "all");
    },
    [beginEditing],
  );

  /**
   * 提交当前编辑。text 省略时从编辑框 DOM 读取（滚动这类没有经过 input 事件的路径）。
   * move 决定提交后选中格往哪走：down/up/right/left 走 moveActive（含边界夹取 + ensureVisible），
   * none = 原地（失焦/滚动），且不抢焦点。
   */
  const finishEditing = useCallback(
    (text: string | undefined, move: CommitMove) => {
      const current = editingRef.current;
      if (!current) return;
      editingRef.current = null; // 同步清空：一次编辑只提交一次（Enter 之后紧跟的 blur 会被挡掉）
      const value = text ?? editorInputRef.current?.value ?? current.text;
      setEditing(null);
      // 编辑结束：拾取模式与拾取状态一起清掉，避免下次编辑带着上一次的虚线框
      formulaModeRef.current = false;
      pickStartRef.current = null;
      setPicking(false);
      callbacksRef.current.onCellCommit?.(current.row, current.col, value);
      if (move === "none") return;
      scrollerRef.current?.focus({ preventScroll: true }); // 收回键盘焦点，方向键可以接着导航
      moveActive(current.row + (move === "down" ? 1 : move === "up" ? -1 : 0), current.col + (move === "right" ? 1 : move === "left" ? -1 : 0));
    },
    [moveActive],
  );
  finishEditingRef.current = finishEditing;

  /** 取消编辑：丢弃本次输入（不回调 onCellCommit），选中格保持不变 */
  const cancelEditing = useCallback(() => {
    if (!editingRef.current) return;
    editingRef.current = null;
    setEditing(null);
    scrollerRef.current?.focus({ preventScroll: true }); // 保持键盘操作连续
  }, []);

  // 编辑能力被关掉（父组件切回只读）或数据变空时，放弃正在进行的编辑，避免留下悬空编辑框
  useEffect(() => {
    if (!editable || empty) cancelEditing();
  }, [cancelEditing, editable, empty]);

  // 编辑过程中父组件把选中移到别处（外部定位/跳转）：先提交当前编辑，再让新的选中生效
  useEffect(() => {
    const current = editingRef.current;
    if (!current || activeCell === undefined || activeCell === null) return;
    if (activeCell.row !== current.row || activeCell.col !== current.col) finishEditing(undefined, "none");
  }, [activeCell, finishEditing]);

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      // 编辑中：按键全部归 input 自己管（方向键移动光标，Enter/Tab/Esc 在编辑框里提交或取消），
      // 网格的键盘导航必须完全让路
      if (editingRef.current) return;
      if (empty) return;
      const current = active ?? { row: 0, col: 0 };
      /** 当前选区：没有显式选区时退化成活动格的 1×1（Ctrl+C / Delete 对单格同样有效） */
      const currentRange = rangeRef.current ?? singleRange(current.row, current.col);
      const mod = event.ctrlKey || event.metaKey;

      // Ctrl+A：整张表（只读下也能用，纯查看能力）
      if (mod && (event.key === "a" || event.key === "A")) {
        event.preventDefault();
        selectAll(true);
        return;
      }

      // Ctrl+C / Ctrl+X：复制（父组件负责序列化与写剪贴板）；Ctrl+X 顺带清空
      if (mod && (event.key === "c" || event.key === "C" || event.key === "x" || event.key === "X")) {
        event.preventDefault();
        const copy = callbacksRef.current.onCopySelection;
        if (copy) copy(currentRange);
        else copyRangeFallback(currentRange); // 兼容旧调用方：网格自己把选区写成 TSV
        if (event.key === "x" || event.key === "X") callbacksRef.current.onClearSelection?.(currentRange);
        return;
      }

      // Ctrl+V：读剪贴板后交给父组件解析；读不到（无权限 / 非安全上下文）静默忽略
      if (mod && (event.key === "v" || event.key === "V")) {
        event.preventDefault();
        const clipboard = navigator.clipboard;
        if (!clipboard?.readText) return;
        const target = currentRange;
        try {
          void clipboard
            .readText()
            .then((text) => {
              if (text) callbacksRef.current.onPaste?.(target, text);
            })
            .catch(() => {});
        } catch {
          /* 忽略剪贴板异常 */
        }
        return;
      }

      // Delete / Backspace：清空选区（是否真写由父组件决定，只读下也照样回调）
      if (event.key === "Delete" || event.key === "Backspace") {
        if (!rangeRef.current && !active) return; // 连活动格都没有：不拦截
        event.preventDefault();
        callbacksRef.current.onClearSelection?.(currentRange);
        return;
      }

      if (editable) {
        // Excel 手感：F2 / Enter 进入编辑且光标落在末尾
        if (event.key === "F2" || event.key === "Enter") {
          event.preventDefault();
          beginEditing(current.row, current.col, "end");
          return;
        }
        // 直接敲可打印字符：该字符成为编辑框初始内容（替换原值）。
        // event.key 长度为 1 且不小于空格即可打印字符；排除 Ctrl/Alt 组合（Ctrl+C 等上面已处理），
        // 输入法组字过程中的 key === "Process" 长度不为 1，因此不会被误判
        if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey && event.key >= " ") {
          event.preventDefault();
          beginEditing(current.row, current.col, "end", event.key);
          return;
        }
      }

      // 翻页步长：一屏可见行数再留 1 行重叠，与 Excel 手感一致
      // （可变行高下按 ROW_H 估算是保守的：换行后一屏放不满这么多行，翻页偏大但与「一屏」同量级）
      const pageSize = Math.max(1, Math.floor(Math.max(0, viewport.h - stickyH) / ROW_H) - 1);
      // Shift 是「以锚点扩展选区」，不带 Shift 则把选区收缩到新落点
      const extend = event.shiftKey;
      switch (event.key) {
        case "ArrowUp": moveActive(current.row - 1, current.col, extend); break;
        case "ArrowDown": moveActive(current.row + 1, current.col, extend); break;
        case "ArrowLeft": moveActive(current.row, current.col - 1, extend); break;
        case "ArrowRight": moveActive(current.row, current.col + 1, extend); break;
        case "Home": moveActive(current.row, 0, extend); break;
        case "End": moveActive(current.row, cols - 1, extend); break;
        case "PageUp": moveActive(current.row - pageSize, current.col, extend); break;
        case "PageDown": moveActive(current.row + pageSize, current.col, extend); break;
        default: return; // 其它按键不拦截
      }
      event.preventDefault();
    },
    [active, beginEditing, cols, copyRangeFallback, editable, empty, moveActive, selectAll, stickyH, viewport.h],
  );

  /**
   * 待提交编辑的「行级切片」。
   * pendingCells 是 `行,列` → 文本 的整表 Map，父组件每改一格都会换一个新的 Map 引用；
   * 若把整张 Map 交给 memo 过的 SheetRow，任何一格改动都会让所有可见行重渲染。
   * 于是这里按行切成 `列 → 文本` 的小切片，并且**内容没变的行复用上一次的切片对象**：
   * SheetRow 的浅比较只对「真正被改过的那一行」失效，其余行的 props 引用不变，直接跳过渲染。
   * 只读（editable=false）时直接返回空表，既不解析 Map 也不产生任何切片（只读路径零开销）。
   */
  const pendingSliceRef = useRef<{ byRow: Map<number, RowPending>; sigs: Map<number, string> }>({
    byRow: new Map(),
    sigs: new Map(),
  });
  const pendingRowSlices = useMemo(() => {
    if (!editable || !pendingCells || pendingCells.size === 0) {
      if (pendingSliceRef.current.byRow.size > 0) pendingSliceRef.current = { byRow: new Map(), sigs: new Map() };
      return pendingSliceRef.current.byRow;
    }
    const prev = pendingSliceRef.current;
    const grouped = new Map<number, Map<number, string>>();
    for (const [key, text] of pendingCells) {
      const sep = key.indexOf(",");
      if (sep <= 0) continue; // 非法 key 直接忽略，不影响渲染
      const row = Number(key.slice(0, sep));
      const col = Number(key.slice(sep + 1));
      if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
      let cols = grouped.get(row);
      if (!cols) {
        cols = new Map<number, string>();
        grouped.set(row, cols);
      }
      cols.set(col, text);
    }
    const byRow = new Map<number, RowPending>();
    const sigs = new Map<number, string>();
    for (const [row, cols] of grouped) {
      // 行内容签名：与上一次相同就复用旧切片对象，这是「只重渲染受影响的行」的关键
      let sig = "";
      for (const [col, text] of cols) sig += `${col}\u0000${text}\u0001`;
      sigs.set(row, sig);
      const cached = prev.byRow.get(row);
      byRow.set(row, cached && prev.sigs.get(row) === sig ? cached : cols);
    }
    pendingSliceRef.current = { byRow, sigs };
    return byRow;
  }, [editable, pendingCells]);

  /**
   * 读最新的公式补全匹配器：走 callbacksRef，所以父组件即使每次传新函数，
   * editorApi 的引用也不变（否则所有可见行都会白重渲染）。
   */
  const readCompleter = useCallback((): FormulaCompleter | undefined => callbacksRef.current.completeFormula, []);

  /**
   * 行组件共用属性：引用稳定，memo 才能挡住滚动带来的整体重渲染。
   * editor 里都是只读 ref / 稳定回调，因此 pendingCells 变化也不会改变它的引用。
   */
  const editorApi = useMemo<RowEditorApi>(
    () => ({
      editable,
      startEdit: startEditFromCell,
      commitEdit: finishEditing,
      cancelEdit: cancelEditing,
      bindInput: bindEditorInput,
      readCompleter,
      bindApi: bindEditorApi,
      readPicking,
      cancelPick,
      notifyFormulaMode,
    }),
    [
      bindEditorApi,
      bindEditorInput,
      cancelEditing,
      cancelPick,
      editable,
      finishEditing,
      notifyFormulaMode,
      readCompleter,
      readPicking,
      startEditFromCell,
    ],
  );

  const sharedRowProps = useMemo(
    () => ({
      colOffsets,
      colWidths,
      getCell,
      onCellMouseDown: handleCellMouseDown,
      onRowMouseDown: handleRowMouseDown,
      onCellContextMenu: handleCellContextMenu,
      onRowContextMenu: handleRowContextMenu,
      editor: editorApi,
      picking,
      onRowResizeStart: beginResizeRow,
      onRowResizeAuto: autoFitRow,
      frozenCols: freezeColCount,
      frozenWidth,
      showColDivider,
      hitKeys,
      currentHitKey,
      onFillHandleDown: beginFillDrag,
    }),
    [
      colOffsets,
      colWidths,
      getCell,
      handleCellMouseDown,
      handleRowMouseDown,
      handleCellContextMenu,
      handleRowContextMenu,
      editorApi,
      picking,
      beginResizeRow,
      autoFitRow,
      freezeColCount,
      frozenWidth,
      showColDivider,
      hitKeys,
      currentHitKey,
      beginFillDrag,
    ],
  );

  /** 选区是否不止一格：多格选区里的活动格不填充（Excel 语义，见 SheetRow 的 multiSelection） */
  const multiSelection = sel !== null && (sel.endRow > sel.startRow || sel.endCol > sel.startCol);
  /** 活动格所在行（-1 = 没有活动格）；只有这一行需要知道「选区是否多格」 */
  const activeRow = active?.row ?? -1;
  /**
   * 选区切片允许覆盖到的最左列：冻结列永远渲染（不随横向滚动走），所以有冻结列时切片必须从第 0 列起，
   * 否则选中「冻结列 + 后面几列」时冻结那一段的外框不会画（SheetRow 再按冻结边界把分段劈开）。
   */
  const sliceFirstCol = freezeColCount > 0 ? 0 : viewRange.firstCol;

  /**
   * 选区在某一行的渲染切片。刻意在这里就把列区间夹到**已渲染的列窗口**：
   * 1) 行组件不用再自己夹，逻辑只有一处；
   * 2) 横向拖出窗口时端点值不变，选区内所有行都不会白重渲染（memo 浅比较全等即跳过）。
   * 返回值全是原始值，直接铺成 SheetRow 的 props。
   */
  const selectionSliceOf = (row: number) => {
    const inSel = sel !== null && row >= sel.startRow && row <= sel.endRow;
    const visStart = inSel ? Math.max(sel.startCol, sliceFirstCol) : -1;
    const visEnd = inSel ? Math.min(sel.endCol, viewRange.lastCol) : -1;
    const show = inSel && visStart <= visEnd;
    return {
      selStartCol: show ? visStart : -1,
      selEndCol: show ? visEnd : -1,
      selTop: inSel && row === sel.startRow,
      selBottom: inSel && row === sel.endRow,
      selLeftEdge: show && sel.startCol >= sliceFirstCol,
      selRightEdge: show && sel.endCol <= viewRange.lastCol,
      // 只有活动格那一行在乎多格与否（其余行的多格值恒为 false），
      // 否则 1×1 ↔ 多格切换时所有行都会因为这一个 props 变化而重渲染
      multiSelection: multiSelection && row === activeRow,
    };
  };

  /* -------------------------------- 渲染 -------------------------------- */

  if (empty) {
    return <div className="flex h-full w-full items-center justify-center bg-app text-[12px] text-faint">空工作表</div>;
  }

  return (
    <div className="relative h-full w-full overflow-hidden bg-app">
      <div
        ref={scrollerRef} tabIndex={0} role="grid" aria-label="表格" aria-rowcount={rows} aria-colcount={cols}
        onScroll={handleScroll} onKeyDown={handleKeyDown}
        /*
         * 这里保持 outline-none（class 不能改成 important 版本：只读与未进入编辑时的 DOM
         * 必须与改动前逐字节一致，test.mjs / edit.mjs 有硬断言）；
         * 真正压住全局 `*:focus-visible` 的那一行在下面的 useLayoutEffect 里用行内样式补。
         */
        className="relative h-full w-full overflow-auto bg-app outline-none"
      >
        {/* 列字母表头：整块 sticky top 冻结。冻结列（freezeCols）的头 + 左上角「全选」放进一个
            行内 sticky 块，横向滚动时整块贴住左边（与数据行的冻结块同一套机制） */}
        <div className="sticky top-0 z-20 border-b border-line bg-panel" style={{ width: contentWidth, height: HEADER_H }}>
          {freezeColCount > 0 ? (
            <div className="sticky left-0 z-30 bg-panel" style={{ width: frozenWidth, height: HEADER_H }}>
              {frozenColList.map((col) => (
                <div
                  key={col} role="columnheader" aria-colindex={col + 1}
                  onMouseDown={(event) => {
                    event.preventDefault();
                    handleColumnMouseDown(col, event.shiftKey, event.clientX, event.clientY);
                  }}
                  onContextMenu={(event) => {
                    if (handleColumnContextMenu(col, event.clientX, event.clientY)) event.preventDefault();
                  }}
                  className={`absolute top-0 truncate border-r border-line text-center text-[11px] ${
                    (active && active.col === col) || (sel !== null && col >= sel.startCol && col <= sel.endCol)
                      ? "bg-accent-soft text-accent"
                      : "bg-panel text-muted"
                  }`}
                  style={{ left: colOffsets[col], width: colWidths[col], height: HEADER_H, lineHeight: `${HEADER_H}px` }}
                >
                  {columnLabel(col)}
                </div>
              ))}
              {/* 左上角：点击 = 全选（等价 Ctrl+A） */}
              <div
                title="全选"
                onMouseDown={handleCornerMouseDown}
                className="absolute left-0 top-0 z-30 cursor-pointer border-r border-line bg-panel"
                style={{ width: GUTTER_W, height: HEADER_H }}
              />
              {frozenColList.map((col) => (
                <div
                  key={`resize-${col}`}
                  role="separator"
                  aria-label={`调整列宽 ${columnLabel(col)}`}
                  aria-orientation="vertical"
                  title="拖动调整列宽，双击按内容自动调整"
                  onMouseDown={(event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    beginResize("col", col, event.clientX, event.clientY);
                  }}
                  onDoubleClick={(event) => {
                    event.stopPropagation();
                    autoFitColumn(col);
                  }}
                  onContextMenu={(event) => event.preventDefault()}
                  className="absolute top-0 z-30 cursor-col-resize hover:bg-accent/40"
                  style={{ left: colOffsets[col] + colWidths[col] - Math.floor(RESIZE_HIT / 2), width: RESIZE_HIT, height: HEADER_H }}
                />
              ))}
              {showColDivider ? (
                <div
                  aria-hidden="true"
                  data-freeze-divider="col"
                  className="pointer-events-none absolute top-0 h-full w-[2px] bg-accent/40"
                  style={{ left: Math.max(0, frozenWidth - 2) }}
                />
              ) : null}
            </div>
          ) : null}
          {visibleCols.filter((col) => col >= freezeColCount).map((col) => (
            <div
              key={col} role="columnheader" aria-colindex={col + 1}
              onMouseDown={(event) => {
                event.preventDefault();
                handleColumnMouseDown(col, event.shiftKey, event.clientX, event.clientY);
              }}
              onContextMenu={(event) => {
                // 父组件接了右键回调才拦系统菜单（没接就保留浏览器默认菜单）
                if (handleColumnContextMenu(col, event.clientX, event.clientY)) event.preventDefault();
              }}
              className={`absolute top-0 truncate border-r border-line text-center text-[11px] ${
                (active && active.col === col) || (sel !== null && col >= sel.startCol && col <= sel.endCol)
                  ? "bg-accent-soft text-accent"
                  : "bg-panel text-muted"
              }`}
              style={{ left: colOffsets[col], width: colWidths[col], height: HEADER_H, lineHeight: `${HEADER_H}px` }}
            >
              {columnLabel(col)}
            </div>
          ))}
          {freezeColCount === 0 ? (
            /* 左上角：点击 = 全选（等价 Ctrl+A） */
            <div
              title="全选"
              onMouseDown={handleCornerMouseDown}
              className="sticky left-0 z-30 cursor-pointer border-r border-line bg-panel"
              style={{ width: GUTTER_W, height: HEADER_H }}
            />
          ) : null}
          {/* 列宽调整命中区：贴在每列右边界上，拖动改宽度、双击按内容自动调整。
              只读网格同样可用（只是看，不改文件内容）；拖动期间 stopPropagation 掉选列逻辑 */}
          {visibleCols.filter((col) => col >= freezeColCount).map((col) => (
            <div
              key={`resize-${col}`}
              role="separator"
              aria-label={`调整列宽 ${columnLabel(col)}`}
              aria-orientation="vertical"
              title="拖动调整列宽，双击按内容自动调整"
              onMouseDown={(event) => {
                event.preventDefault();
                event.stopPropagation();
                beginResize("col", col, event.clientX, event.clientY);
              }}
              onDoubleClick={(event) => {
                event.stopPropagation();
                autoFitColumn(col);
              }}
              onContextMenu={(event) => event.preventDefault()}
              className="absolute top-0 z-30 cursor-col-resize hover:bg-accent/40"
              style={{ left: colOffsets[col] + colWidths[col] - Math.floor(RESIZE_HIT / 2), width: RESIZE_HIT, height: HEADER_H }}
            />
          ))}
        </div>

        {/* 冻结行（freezeRows / headerRow 同一套）：占流式高度并 sticky 在列字母表头下方，
            多行冻结时按各自高度依次堆叠（sticky top 逐行累加），其余行照常在滚动区里 */}
        {Array.from({ length: empty ? 0 : headerOffset }, (_, row) => {
          let top = HEADER_H;
          for (let i = 0; i < row; i += 1) top += rowIndex.heightOf(i);
          return (
            <SheetRow
              {...sharedRowProps}
              key={`frozen-${row}`}
              row={row} label={String(row + 1)} top={top} width={contentWidth} frozen
              panel={row === headerStyledRow}
              showRowDivider={showRowDivider && row === headerOffset - 1}
              height={rowIndex.heightOf(row)} wrap={wrapEnabled}
              firstCol={viewRange.firstCol} lastCol={viewRange.lastCol}
              activeCol={active && active.row === row ? active.col : -1}
              pendingRow={pendingRowSlices.get(row)}
              editingCol={editing !== null && editing.row === row ? editing.col : -1}
              editingText={editing !== null && editing.row === row ? editing.text : ""}
              editingCaret={editing !== null && editing.row === row ? editing.caret : "end"}
              fillHandleCol={fillCorner && fillCorner.row === row ? fillCorner.col : -1}
              {...selectionSliceOf(row)}
            />
          );
        })}

        {/* 滚动区：高度 = 各行高度之和（行高索引给出的内容总高）；行元素按 top 绝对定位。
            背景画一层淡行线：快速拖动时尚未渲染的区域显示为「空表格」而不是一片白。
            **只有全表都是默认行高时才画**：这层纹理是固定 24px 间距的，行高一变（换行折行、
            手动调整行高）就会与真实行边界错位（用户反馈的「调整行高后布满横线」）。
            有非默认行高时宁可不画这层，行边界由每行自己的 border-b（画在真实行高处）给出 ——
            内容区以下也不会多出与真实行边界不符的线（这里的高度就是行高索引的总高）。 */}
        <div
          className="relative"
          style={{
            width: contentWidth,
            height: rowIndex.totalHeight,
            contain: "layout paint",
            backgroundImage: uniformRowHeight
              ? `repeating-linear-gradient(to bottom, var(--color-line) 0 1px, transparent 1px ${ROW_H}px)`
              : undefined,
          }}
        >
          {/*
            公式引用框：放在行之前（所以在单元格底色之下、在行号栏与冻结表头之下），
            只画虚线边框、pointer-events-none —— 不改变选区样式，也不拦任何鼠标事件。
          */}
          {formulaFrames.length > 0 ? (
            <div aria-hidden="true" className="pointer-events-none absolute left-0 top-0">
              {formulaFrames.map((frame) => (
                <div
                  key={frame.key}
                  data-formula-frame="true"
                  className={`absolute border-dashed ${frame.tone}`}
                  style={{
                    left: frame.left,
                    top: frame.top,
                    width: frame.width,
                    height: frame.height,
                    borderTopWidth: frame.showTop ? 2 : 0,
                    borderRightWidth: frame.showRight ? 2 : 0,
                    borderBottomWidth: frame.showBottom ? 2 : 0,
                    borderLeftWidth: frame.showLeft ? 2 : 0,
                  }}
                />
              ))}
            </div>
          ) : null}
          {visibleRows.map((row) => (
            <SheetRow
              {...sharedRowProps}
              key={row}
              row={row} label={String(row + 1)} top={rowIndex.topOf(row)} width={contentWidth} frozen={false}
              panel={false} showRowDivider={false}
              height={rowIndex.heightOf(row)} wrap={wrapEnabled}
              firstCol={viewRange.firstCol} lastCol={viewRange.lastCol}
              activeCol={active && active.row === row ? active.col : -1}
              /* 行级切片 + 只属于本行的编辑态与选区：无关行的这些 props 全等，memo 直接跳过 */
              pendingRow={pendingRowSlices.get(row)}
              editingCol={editing !== null && editing.row === row ? editing.col : -1}
              editingText={editing !== null && editing.row === row ? editing.text : ""}
              editingCaret={editing !== null && editing.row === row ? editing.caret : "end"}
              fillHandleCol={fillCorner && fillCorner.row === row ? fillCorner.col : -1}
              {...selectionSliceOf(row)}
            />
          ))}
          {/* 拖拽填充的虚线预览：从原选区延伸到指针所在的整行/整列（只扩一个方向）。
              放在行之后、pointer-events-none —— 只看不拦，也不改选区与内容。 */}
          {fillPreview && !sameRange(fillPreview.source, fillPreview.target)
            ? (() => {
                const target = fillPreview.target;
                const lastCol = Math.max(0, Math.min(cols - 1, target.endCol));
                const firstCol = Math.max(0, Math.min(cols - 1, target.startCol));
                // 冻结行不在内容坐标系里（topOf 会是负数），预览从正文第一行起算
                const firstRow = Math.max(headerOffset, Math.min(rows - 1, target.startRow));
                const lastRow = Math.max(headerOffset, Math.min(rows - 1, target.endRow));
                const left = colOffsets[firstCol];
                const right = colOffsets[lastCol] + colWidths[lastCol];
                const top = Math.max(0, rowIndex.topOf(firstRow));
                const bottom = rowIndex.topOf(lastRow) + rowIndex.heightOf(lastRow);
                return (
                  <div
                    aria-hidden="true"
                    data-fill-preview="true"
                    className="pointer-events-none absolute z-[2] border-2 border-dashed border-accent"
                    style={{
                      left,
                      top,
                      width: Math.max(1, right - left),
                      height: Math.max(1, bottom - top),
                    }}
                  />
                );
              })()
            : null}
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

      {/* 拖动调整宽高时的落点指示线（只在拖动期间存在） */}
      {resizeGuide ? (
        <div
          aria-hidden="true"
          className={`pointer-events-none absolute z-40 bg-accent ${
            resizeGuide.kind === "col" ? "top-0 h-full w-px" : "left-0 h-px w-full"
          }`}
          style={resizeGuide.kind === "col" ? { left: resizeGuide.client } : { top: resizeGuide.client }}
        />
      ) : null}
    </div>
  );
}
