/**
 * 菜单栏 → 表格视图的一次性命令通道。
 *
 * 工具栏是全局组件，而"自动调整表格""插入公式"这类动作要落到当前打开的表格视图里
 * （它持有选区、活动单元格、列宽行高等组件内状态）。用模块级的小型事件通道比把这些状态
 * 全搬进全局 store 更轻，也避免工具栏与表格视图互相依赖。
 *
 * 只有**当前激活**的表格视图会响应（见 SheetView 里的 isActiveDoc 判断），
 * 分屏时不会两个视图同时执行同一条命令。
 */

export type SheetCommand =
  | { kind: "autoFit"; target?: "columns" | "rows" | "both" }
  | { kind: "insertFormula"; template: string }
  | { kind: "freeze"; rows: number; cols: number };

type Listener = (command: SheetCommand) => void;

const listeners = new Set<Listener>();

/** 发送一条命令（工具栏调用） */
export function sendSheetCommand(command: SheetCommand): void {
  for (const listener of listeners) {
    try {
      listener(command);
    } catch (error) {
      console.error("[MasterEdit] 表格命令执行失败", command.kind, error);
    }
  }
}

/** 订阅命令（表格视图调用），返回取消订阅函数 */
export function onSheetCommand(listener: Listener): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** 常用公式模板：`{range}` 会被替换成当前选区的引用（没有选区时用活动单元格） */
export const COMMON_FORMULA_TEMPLATES: Array<{ label: string; template: string }> = [
  { label: "求和 SUM", template: "=SUM({range})" },
  { label: "平均值 AVERAGE", template: "=AVERAGE({range})" },
  { label: "计数 COUNT", template: "=COUNT({range})" },
  { label: "最大值 MAX", template: "=MAX({range})" },
  { label: "最小值 MIN", template: "=MIN({range})" },
  { label: "四舍五入 ROUND", template: "=ROUND({range},2)" },
  { label: "条件判断 IF", template: '=IF({range}>0,"是","否")' },
  { label: "文本长度 LEN", template: "=LEN({range})" },
  { label: "纵向查找 VLOOKUP", template: "=VLOOKUP({range},A:C,3,0)" },
];
