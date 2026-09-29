/**
 * 表格查找 / 替换的文本处理（CSV 与 xlsx 共用）。
 *
 * 这里只做"单元格文本 ↔ 查询词"的匹配与替换，不关心数据从哪来：
 * - CSV 的查找直接用 `buildCellMatcher` 在内存表格上扫描；
 * - xlsx 的查找由后端 `spreadsheet_find` 完成（只支持区分大小写与全字），
 *   但**替换**需要在前端把命中格的新文本算出来，因此同样走 `replaceCellText`。
 *
 * 两条约定：
 * 1. 查询词默认按**字面量**处理，只有打开 `regex` 才当正则（xlsx 不支持正则，面板会置灰）；
 * 2. 替换串里的 `$` 必须转义 —— JS 的 `String.replace` 会把它当"捕获组引用"，
 *    用户输入 `$1` 会被替换成空，属于静默改错内容。
 */

export interface SheetReplaceOptions {
  /** 区分大小写 */
  matchCase: boolean;
  /** 全字匹配：整格内容完全相同才算命中 */
  wholeCell: boolean;
  /** 把查询词当正则表达式 */
  regex: boolean;
}

/** 把查询词编译成正则（字面量时转义元字符）；返回 null 表示正则非法 */
export function compileQuery(query: string, options: SheetReplaceOptions): RegExp | null {
  if (query === "") return null;
  const source = options.regex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  try {
    return new RegExp(options.wholeCell ? `^(?:${source})$` : source, options.matchCase ? "" : "i");
  } catch {
    return null;
  }
}

/** 单元格文本是否命中查询（CSV 查找用；正则非法时返回 false，由调用方给出错误提示） */
export function buildCellMatcher(
  query: string,
  options: SheetReplaceOptions,
): ((text: string) => boolean) | null {
  const regex = compileQuery(query, options);
  if (!regex) return null;
  return (text: string) => regex.test(text);
}

/** 替换串转义：避免 `$1`、`$&` 被当成捕获组引用 */
function escapeReplacement(replacement: string): string {
  return replacement.replace(/\$/g, "$$$$");
}

/**
 * 把单元格文本里的查询词替换掉。
 * - `all = false`：只替换**第一处**（对应面板的「替换」）
 * - `all = true`：替换该格内**全部**出现（对应「全部替换」）
 * - 没有命中或结果与原文相同 → 返回 null（调用方据此跳过，避免产生无意义的编辑）
 */
export function replaceCellText(
  text: string,
  query: string,
  replacement: string,
  options: SheetReplaceOptions,
  all: boolean,
): string | null {
  const regex = compileQuery(query, options);
  if (!regex) return null;
  if (options.wholeCell) {
    if (!regex.test(text)) return null;
    return replacement === text ? null : replacement;
  }
  if (!regex.test(text)) return null;
  // test 之后重新构造（带 g 的正则 test 会推进 lastIndex，直接复用会漏掉第一处）
  const pattern = options.regex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const flags = `${options.matchCase ? "" : "i"}${all ? "g" : ""}`;
  const next = text.replace(new RegExp(pattern, flags), escapeReplacement(replacement));
  return next === text ? null : next;
}

/** 判断某个单元格是否落在选区里（范围筛选用，坐标为 0 起、含端点） */
export function inSelectionRange(
  range: { startRow: number; startCol: number; endRow: number; endCol: number } | null,
  row: number,
  col: number,
): boolean {
  if (!range) return true;
  return (
    row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
  );
}
