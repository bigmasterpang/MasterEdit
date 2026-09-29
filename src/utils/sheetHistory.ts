/**
 * 表格编辑的撤销 / 重做历史（单元格编辑 + 结构操作统一成一条栈）。
 *
 * 为什么放在模块级 Map 而不是组件内的 ref：
 * 表格视图在切换标签、切换视图（表格 / 源码 / 分屏）、切换工作表时都可能卸载重建，
 * 组件内的 ref 会随之丢失，表现为「改完保存后过一会儿 Ctrl+Z 就没反应了」。
 * 这里按文档 id 保存，只要文档还开着历史就一直有效；文档关闭时由 clearSheetHistory 清理。
 *
 * 两种步骤：
 * - `cells`：单元格编辑（含清空选区、粘贴），值保存在前端，撤销即把值改回去；
 * - `structure`：插入/删除行列、工作表增删改复制 —— 这些改动在 Rust 侧的内存影子里，
 *   撤销由后端快照栈完成（`spreadsheet_undo` / `spreadsheet_redo`）。
 *
 * ## 坐标为什么会错位，以及这里怎么解决
 *
 * 插入/删除行列会平移所有单元格的坐标，而**结构操作本身也可以被撤销**（撤销后坐标又平移回来）。
 * 因此不能"就地改写历史里的行号"——那样撤销结构操作后，旧步骤的坐标就再也回不去了
 * （实测症状：删掉第 3 列后撤销恢复该列，再撤销时本该清空第 3 列的编辑，却清空了第 4 列的）。
 *
 * 这里的做法是**布局版本链**：每发生一次结构操作就往 `layoutOps` 追加一条变换；每个单元格步骤
 * 记下自己产生时的版本号；应用步骤时把坐标**按之后的变换依次换算**。结构操作被撤销时把链尾弹出，
 * 于是被删除区域里的步骤会自动"重新生效"，而不会算错格子。
 */

/** 一步编辑里的一格：key 为 `工作表|行|列`（CSV 用 `行,列`），before/after 为 null 表示「无编辑」 */
export interface HistoryCell<T> {
  key: string;
  before: T | null;
  after: T | null;
}

/** 单元格编辑步骤：可以只含一格，也可以是一批（清空选区 / 粘贴） */
export interface CellStep<T> {
  kind: "cells";
  cells: Array<HistoryCell<T>>;
  /** 产生这一步时的布局版本（= 当时的 layoutOps 长度） */
  layoutVersion: number;
}

/** 结构操作的类型（与后端 StructureOp 的 kind 一致） */
export type StructureOpKind =
  | "insertRows"
  | "deleteRows"
  | "insertCols"
  | "deleteCols"
  | "addSheet"
  | "copySheet"
  | "renameSheet"
  | "deleteSheet";

/** 会平移坐标的结构操作描述 */
export type LayoutOp =
  | { kind: "insertRows" | "deleteRows" | "insertCols" | "deleteCols"; sheet: string; at: number; count: number }
  | { kind: "renameSheet"; sheet: string; name: string }
  | { kind: "deleteSheet"; sheet: string };

/** 结构操作步骤（实际状态在 Rust 影子的快照栈里） */
export interface StructureStep {
  kind: "structure";
  /** 用于提示文案，例如「插入 2 行」 */
  label: string;
  /** 这次结构操作对坐标的影响（撤销时从布局链尾弹出） */
  op: LayoutOp | null;
}

/**
 * 布局调整步骤（手动改列宽 / 行高）。
 * key 形如 `col:3` / `row:5`；before/after 是像素值，`null` 表示「未手动设置（跟随自动/采样）」。
 */
export interface LayoutStep {
  kind: "layout";
  changes: Array<{ key: string; before: number | null; after: number | null }>;
  /** 产生这一步时的布局版本（与单元格步骤同理，应用时按之后的变换换算） */
  layoutVersion: number;
}

/**
 * 内容步骤（CSV / TSV 的行列结构操作）：整体改写文本，一步还原。
 * CSV 没有工作簿结构，插入删除行列最终都是文本改写，用前后文本记录最稳妥。
 */
export interface ContentStep {
  kind: "content";
  before: string;
  after: string;
}

export type HistoryStep<T> = CellStep<T> | StructureStep | LayoutStep | ContentStep;

/** 撤销方向 */
export type HistoryDirection = "undo" | "redo";

interface HistorySlot {
  undo: unknown[];
  redo: unknown[];
  /** 布局版本链：按时间顺序记录会平移坐标的结构操作 */
  layoutOps: LayoutOp[];
}

const histories = new Map<string, HistorySlot>();

function slotOf(docId: string): HistorySlot {
  let slot = histories.get(docId);
  if (!slot) {
    slot = { undo: [], redo: [], layoutOps: [] };
    histories.set(docId, slot);
  }
  return slot;
}

/** 取某个文档的历史容器（不存在则创建） */
export function sheetHistory<T>(docId: string): {
  undo: Array<HistoryStep<T>>;
  redo: Array<HistoryStep<T>>;
} {
  return slotOf(docId) as unknown as { undo: Array<HistoryStep<T>>; redo: Array<HistoryStep<T>> };
}


/** 记录一步单元格编辑（任何新操作都会清空重做栈） */
export function pushCellStep<T>(docId: string, cells: Array<HistoryCell<T>>): void {
  const slot = slotOf(docId);
  const step: CellStep<T> = { kind: "cells", cells, layoutVersion: slot.layoutOps.length };
  slot.undo.push(step as unknown);
  slot.redo.length = 0;
}

/**
 * 把一次「连续调整」合并进上一步（同一个 key）。
 *
 * 拖动列宽/行高时网格每帧都会回调一次，若每次都压一步，拖一下就会产生几十步撤销
 * （用户实测「调整行高记录了 80 多个操作」）。拖动会话内只更新最后一步的 after，
 * 松手（window mouseup）后会话结束，下一次拖动才产生新的一步。
 */
export function amendLastLayoutStep(
  docId: string,
  changes: Array<{ key: string; after: number | null }>,
): boolean {
  const slot = slotOf(docId);
  const last = slot.undo[slot.undo.length - 1] as HistoryStep<unknown> | undefined;
  if (!last || last.kind !== "layout") return false;
  let matched = false;
  for (const change of changes) {
    const target = last.changes.find((entry) => entry.key === change.key);
    if (!target) continue;
    target.after = change.after;
    matched = true;
  }
  return matched;
}

/** 记录一步内容改写（CSV 行列结构操作） */
export function pushContentStep(docId: string, before: string, after: string): void {
  if (before === after) return;
  const slot = slotOf(docId);
  slot.undo.push({ kind: "content", before, after } as unknown);
  slot.redo.length = 0;
}

/** 记录一步布局调整（列宽 / 行高），与内容、结构操作共用同一条撤销栈 */
export function pushLayoutStep(
  docId: string,
  changes: Array<{ key: string; before: number | null; after: number | null }>,
): void {
  if (changes.length === 0) return;
  const slot = slotOf(docId);
  const step: LayoutStep = { kind: "layout", changes, layoutVersion: slot.layoutOps.length };
  slot.undo.push(step as unknown);
  slot.redo.length = 0;
}

/** 记录一步结构操作：同时把它对坐标的影响追加到布局链 */
export function pushStructureStep(docId: string, label: string, op: LayoutOp | null): void {
  const slot = slotOf(docId);
  if (op) slot.layoutOps.push(op);
  const step: StructureStep = { kind: "structure", label, op };
  slot.undo.push(step as unknown);
  slot.redo.length = 0;
}

/** 结构操作被撤销：把它的坐标影响从链尾弹出 */
export function popLayoutOp(docId: string, op: LayoutOp | null): void {
  if (!op) return;
  const slot = slotOf(docId);
  const last = slot.layoutOps[slot.layoutOps.length - 1];
  if (last === op) slot.layoutOps.pop();
}

/** 结构操作被重做：把坐标影响重新压回链尾 */
export function pushLayoutOpAgain(docId: string, op: LayoutOp | null): void {
  if (!op) return;
  slotOf(docId).layoutOps.push(op);
}

/**
 * 把某个版本下的单元格 key 换算到当前布局。
 * 返回 null 表示该格在当前布局里不存在（例如所在行/列已被删除）——此时跳过这一格。
 */
export function mapCellKey(docId: string, key: string, layoutVersion: number): string | null {
  const slot = slotOf(docId);
  const parts = key.split("|");
  // 布局步骤的 key（`col:3` / `row:5`）：跟着行列平移，落在被删行列里则丢弃
  if (parts.length === 1 && /^(col|row):\d+$/.test(key)) {
    const [kind, indexText] = key.split(":");
    let index = Number(indexText);
    const isRow = kind === "row";
    for (let at = layoutVersion; at < slot.layoutOps.length; at += 1) {
      const op = slot.layoutOps[at];
      if (op.kind === "renameSheet" || op.kind === "deleteSheet") continue;
      const opIsRow = op.kind === "insertRows" || op.kind === "deleteRows";
      if (opIsRow !== isRow) continue;
      if (op.kind === "deleteRows" || op.kind === "deleteCols") {
        if (index >= op.at && index < op.at + op.count) return null;
        if (index >= op.at + op.count) index -= op.count;
      } else if (index >= op.at) {
        index += op.count;
      }
    }
    return `${kind}:${index}`;
  }
  // CSV 的 key 是 `行,列`（没有工作表，也不会有结构操作）
  if (parts.length !== 3) return key;
  let [sheetName, rowText, colText] = parts;
  let row = Number(rowText);
  let col = Number(colText);
  for (let index = layoutVersion; index < slot.layoutOps.length; index += 1) {
    const op = slot.layoutOps[index];
    if (op.kind === "renameSheet") {
      if (sheetName === op.sheet) sheetName = op.name;
      continue;
    }
    if (op.kind === "deleteSheet") {
      if (sheetName === op.sheet) return null;
      continue;
    }
    if (sheetName !== op.sheet) continue;
    const isRow = op.kind === "insertRows" || op.kind === "deleteRows";
    const position = isRow ? row : col;
    if (op.kind === "deleteRows" || op.kind === "deleteCols") {
      if (position >= op.at && position < op.at + op.count) return null;
      if (position >= op.at + op.count) {
        if (isRow) row -= op.count;
        else col -= op.count;
      }
    } else if (position >= op.at) {
      if (isRow) row += op.count;
      else col += op.count;
    }
  }
  return `${sheetName}|${row}|${col}`;
}

/** 换算一整批单元格（跳过在当前布局里不存在的格子） */
export function mapCellBatch<T>(
  docId: string,
  cells: Array<HistoryCell<T>>,
  layoutVersion: number,
): Array<HistoryCell<T>> {
  const mapped: Array<HistoryCell<T>> = [];
  for (const cell of cells) {
    const key = mapCellKey(docId, cell.key, layoutVersion);
    if (key === null) continue;
    mapped.push({ ...cell, key });
  }
  return mapped;
}

/**
 * 取出一端的步骤（不移动）。
 * 执行成功后再用 `commitHistoryStep` 移到另一端；失败用 `returnHistoryStep` 放回原处，
 * 保证「撤销失败时栈不变」。
 */
export function popHistoryStep<T>(docId: string, direction: HistoryDirection): HistoryStep<T> | null {
  const slot = slotOf(docId);
  const stack = direction === "undo" ? slot.undo : slot.redo;
  return (stack.pop() as HistoryStep<T> | undefined) ?? null;
}

/** 把步骤移到另一端（撤销成功 → 进重做栈；重做成功 → 进撤销栈） */
export function commitHistoryStep<T>(
  docId: string,
  direction: HistoryDirection,
  step: HistoryStep<T>,
): void {
  const slot = slotOf(docId);
  (direction === "undo" ? slot.redo : slot.undo).push(step as unknown);
}

/** 执行失败：把步骤放回原处 */
export function returnHistoryStep<T>(
  docId: string,
  direction: HistoryDirection,
  step: HistoryStep<T>,
): void {
  const slot = slotOf(docId);
  (direction === "undo" ? slot.undo : slot.redo).push(step as unknown);
}

/** 历史深度（界面按钮的可用状态用） */
export function sheetHistoryDepth(docId: string): { undo: number; redo: number } {
  const slot = histories.get(docId);
  return { undo: slot?.undo.length ?? 0, redo: slot?.redo.length ?? 0 };
}

/** 文档关闭时清理，避免长时间会话里堆积 */
export function clearSheetHistory(docId: string): void {
  histories.delete(docId);
}
