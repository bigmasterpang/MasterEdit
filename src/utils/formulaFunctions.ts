/**
 * Excel 常用函数清单与「输入 = 号后的补全」匹配。
 *
 * 这里只做**名称与用法提示**补全（不解析参数、不计算公式），因此不需要公式引擎：
 * 输入 `=SU` 时提示 SUM / SUMIF / SUMIFS / SUMPRODUCT 等，回车插入 `SUM(`。
 * 中文描述用于提示条，签名按 Excel 的习惯写（[] 表示可选参数）。
 */

export interface FormulaFunction {
  /** 函数名（Excel 中英版本通用英文名）；引用类候选项这里放引用文本，如 `B:B` */
  name: string;
  /** 用法签名，例如 SUM(number1, [number2], …) */
  signature: string;
  /** 一句话说明 */
  description: string;
  /** 无参数函数：补全时直接插入 `()` */
  noArgs?: boolean;
  /**
   * 插入方式：
   * - `call`（默认）：函数，插入 `NAME(`
   * - `plain`：区域引用，原样插入（如 `B:B`）
   */
  insert?: "call" | "plain";
}

/** 常用函数：按使用频率大致排序，空查询时优先展示前面这些 */
export const FORMULA_FUNCTIONS: FormulaFunction[] = [
  { name: "SUM", signature: "SUM(number1, [number2], …)", description: "求和" },
  { name: "AVERAGE", signature: "AVERAGE(number1, [number2], …)", description: "平均值" },
  { name: "COUNT", signature: "COUNT(value1, [value2], …)", description: "统计数字个数" },
  { name: "COUNTA", signature: "COUNTA(value1, [value2], …)", description: "统计非空个数" },
  { name: "COUNTIF", signature: "COUNTIF(range, criteria)", description: "按条件计数" },
  { name: "COUNTIFS", signature: "COUNTIFS(range1, c1, [range2, c2], …)", description: "多条件计数" },
  { name: "SUMIF", signature: "SUMIF(range, criteria, [sum_range])", description: "按条件求和" },
  { name: "SUMIFS", signature: "SUMIFS(sum_range, range1, c1, …)", description: "多条件求和" },
  { name: "AVERAGEIF", signature: "AVERAGEIF(range, criteria, [avg_range])", description: "按条件求平均" },
  { name: "AVERAGEIFS", signature: "AVERAGEIFS(avg_range, range1, c1, …)", description: "多条件平均" },
  { name: "MAX", signature: "MAX(number1, [number2], …)", description: "最大值" },
  { name: "MIN", signature: "MIN(number1, [number2], …)", description: "最小值" },
  { name: "LARGE", signature: "LARGE(array, k)", description: "第 k 大的值" },
  { name: "SMALL", signature: "SMALL(array, k)", description: "第 k 小的值" },
  { name: "MEDIAN", signature: "MEDIAN(number1, [number2], …)", description: "中位数" },
  { name: "RANK", signature: "RANK(number, ref, [order])", description: "排名" },
  { name: "ROUND", signature: "ROUND(number, num_digits)", description: "四舍五入" },
  { name: "ROUNDUP", signature: "ROUNDUP(number, num_digits)", description: "向上舍入" },
  { name: "ROUNDDOWN", signature: "ROUNDDOWN(number, num_digits)", description: "向下舍入" },
  { name: "INT", signature: "INT(number)", description: "取整" },
  { name: "ABS", signature: "ABS(number)", description: "绝对值" },
  { name: "MOD", signature: "MOD(number, divisor)", description: "取余" },
  { name: "POWER", signature: "POWER(number, power)", description: "乘幂" },
  { name: "SQRT", signature: "SQRT(number)", description: "平方根" },
  { name: "IF", signature: "IF(条件, 真时值, [假时值])", description: "条件判断" },
  { name: "IFS", signature: "IFS(条件1, 值1, [条件2, 值2], …)", description: "多条件判断" },
  { name: "IFERROR", signature: "IFERROR(value, value_if_error)", description: "出错时返回备用值" },
  { name: "AND", signature: "AND(条件1, [条件2], …)", description: "全部成立" },
  { name: "OR", signature: "OR(条件1, [条件2], …)", description: "任一成立" },
  { name: "NOT", signature: "NOT(条件)", description: "取反" },
  { name: "VLOOKUP", signature: "VLOOKUP(查找值, 区域, 列号, [匹配方式])", description: "纵向查找" },
  { name: "HLOOKUP", signature: "HLOOKUP(查找值, 区域, 行号, [匹配方式])", description: "横向查找" },
  { name: "XLOOKUP", signature: "XLOOKUP(查找值, 查找区域, 返回区域, [未找到值])", description: "新版查找（推荐）" },
  { name: "INDEX", signature: "INDEX(array, row_num, [col_num])", description: "按位置取值" },
  { name: "MATCH", signature: "MATCH(查找值, 区域, [匹配方式])", description: "返回位置" },
  { name: "OFFSET", signature: "OFFSET(reference, rows, cols, [height], [width])", description: "偏移引用" },
  { name: "INDIRECT", signature: "INDIRECT(文本引用, [a1])", description: "文本转引用" },
  { name: "ROW", signature: "ROW([reference])", description: "当前行号" },
  { name: "COLUMN", signature: "COLUMN([reference])", description: "当前列号" },
  { name: "ROWS", signature: "ROWS(array)", description: "行数" },
  { name: "COLUMNS", signature: "COLUMNS(array)", description: "列数" },
  { name: "SUMPRODUCT", signature: "SUMPRODUCT(array1, [array2], …)", description: "数组相乘后求和" },
  { name: "SUBTOTAL", signature: "SUBTOTAL(函数号, ref1, …)", description: "分类汇总（忽略隐藏行）" },
  { name: "LEN", signature: "LEN(text)", description: "文本长度" },
  { name: "LEFT", signature: "LEFT(text, [num_chars])", description: "取左侧字符" },
  { name: "RIGHT", signature: "RIGHT(text, [num_chars])", description: "取右侧字符" },
  { name: "MID", signature: "MID(text, start_num, num_chars)", description: "取中间字符" },
  { name: "FIND", signature: "FIND(找什么, 在哪找, [起始位置])", description: "查找位置（区分大小写）" },
  { name: "SEARCH", signature: "SEARCH(找什么, 在哪找, [起始位置])", description: "查找位置（不区分大小写）" },
  { name: "SUBSTITUTE", signature: "SUBSTITUTE(text, old, new, [instance])", description: "替换文本" },
  { name: "REPLACE", signature: "REPLACE(old_text, start, num, new_text)", description: "按位置替换" },
  { name: "TRIM", signature: "TRIM(text)", description: "去掉多余空格" },
  { name: "UPPER", signature: "UPPER(text)", description: "转大写" },
  { name: "LOWER", signature: "LOWER(text)", description: "转小写" },
  { name: "TEXT", signature: "TEXT(value, format_text)", description: "按格式转文本" },
  { name: "VALUE", signature: "VALUE(text)", description: "文本转数字" },
  { name: "TEXTJOIN", signature: "TEXTJOIN(分隔符, 忽略空, text1, …)", description: "按分隔符连接" },
  { name: "CONCAT", signature: "CONCAT(text1, [text2], …)", description: "连接文本" },
  { name: "REPT", signature: "REPT(text, number_times)", description: "重复文本" },
  { name: "TODAY", signature: "TODAY()", description: "今天日期", noArgs: true },
  { name: "NOW", signature: "NOW()", description: "当前日期时间", noArgs: true },
  { name: "DATE", signature: "DATE(年, 月, 日)", description: "构造日期" },
  { name: "YEAR", signature: "YEAR(日期)", description: "取年份" },
  { name: "MONTH", signature: "MONTH(日期)", description: "取月份" },
  { name: "DAY", signature: "DAY(日期)", description: "取日" },
  { name: "DAYS", signature: "DAYS(结束日期, 开始日期)", description: "相差天数" },
  { name: "DATEDIF", signature: "DATEDIF(开始, 结束, 单位)", description: "日期间隔（Y/M/D）" },
  { name: "WEEKDAY", signature: "WEEKDAY(日期, [返回类型])", description: "星期几" },
  { name: "EOMONTH", signature: "EOMONTH(开始日期, 月数)", description: "月末日期" },
  { name: "NETWORKDAYS", signature: "NETWORKDAYS(开始, 结束, [节假日])", description: "工作日天数" },
  { name: "ISBLANK", signature: "ISBLANK(value)", description: "是否为空" },
  { name: "ISNUMBER", signature: "ISNUMBER(value)", description: "是否为数字" },
  { name: "ISTEXT", signature: "ISTEXT(value)", description: "是否为文本" },
  { name: "ISERROR", signature: "ISERROR(value)", description: "是否出错" },
  { name: "UNIQUE", signature: "UNIQUE(array)", description: "去重（动态数组）" },
  { name: "SORT", signature: "SORT(array, [sort_index], [order])", description: "排序（动态数组）" },
  { name: "FILTER", signature: "FILTER(array, include, [if_empty])", description: "筛选（动态数组）" },
];

export interface FormulaSuggestion {
  /** 需要替换掉的文本范围（相对整个输入框文本） */
  from: number;
  to: number;
  /** 输入的函数名前缀（用于过滤） */
  query: string;
  items: FormulaFunction[];
}

/** 空查询时最多展示几个函数（输入 `=` 立刻弹出的那一批） */
const EMPTY_QUERY_LIMIT = 10;
/** 有关键字时最多展示几个 */
const QUERY_LIMIT = 10;
/** 超过这么长的函数名还没输完就不再提示（避免在普通文本里乱弹） */
const MAX_QUERY_LENGTH = 20;

/**
 * 根据输入框文本与光标位置算出补全建议。
 *
 * 规则（与 Excel 接近）：
 * - 只在**公式里**提示：`=` 开头（允许前面有空格）；
 * - 取光标前连续的「字母 / 数字 / 点 / 下划线」作为查询词；
 * - 查询词为空时，只有紧跟 `=`（或 `=(`、`+`、`-`、`,`、`(`）才弹出常用函数；
 * - 命中按「前缀优先、其次包含」排序，前缀命中里按清单顺序（≈ 常用度）；
 * - 另外按 `context` 里的已用行列数补上**区域引用**候选（整列 `B:B`、整行 `2:2`、已用区域），
 *   让 `=SUM(B` 能直接选到整列 B —— 用户不必记 Excel 的引用写法。
 */
export function matchFormulaFunctions(
  text: string,
  caret: number,
  context?: { usedRows?: number; usedCols?: number },
): FormulaSuggestion | null {
  const head = text.slice(0, Math.max(0, Math.min(caret, text.length)));
  if (!/^\s*=/.test(text)) return null;

  // 注意必须有捕获组：下面用 token[1] 取「光标前连续的名字字符」，
  // 若写成无捕获组的 /[A-Za-z][A-Za-z0-9._]*$/，token[1] 是 undefined，
  // query.length 会直接抛 TypeError（TypeScript 不会报，RegExpExecArray 的下标访问类型是 string）。
  // 允许以数字开头：这样才能提示「整行引用」（2:2）这类候选。
  const token = /([A-Za-z0-9][A-Za-z0-9._]*)$/.exec(head);
  const query = token ? token[1] : "";
  const from = head.length - query.length;
  if (query.length > MAX_QUERY_LENGTH) return null;

  const usedRows = context?.usedRows ?? 0;
  const usedCols = context?.usedCols ?? 0;
  /**
   * 光标前一个非空字符：判断是否处于「新参数位置」（=、(、,、运算符之后）。
   * 注意要从**查询词之前**取，不能直接取光标前一个字符 —— 否则 `=SUM(2` 取到的是
   * 查询词自己的 "2"，会被误判成「不是新参数位置」而不提示整行引用。
   */
  const previous = head.slice(0, from).replace(/\s+$/, "").slice(-1);
  const freshArgument =
    previous === "=" ||
    previous === "(" ||
    previous === "," ||
    previous === "+" ||
    previous === "-" ||
    previous === "*" ||
    previous === "/";
  /**
   * 更严格的「参数起始位置」：只认 = ( , 三种。
   * 数字查询（整行引用）只在这里提示 —— `=SUM(2` 想要 `2:2`，
   * 而 `=A1+2` 里的 2 是运算数，提示整行引用纯属噪音。
   * 字母查询（整列引用）不受此限制：`=A1+B` 时提示 B 列正是用户想要的。
   */
  const argumentStart = previous === "=" || previous === "(" || previous === ",";

  if (query === "") {
    // 只有紧跟在新参数位置才提示，避免在 `=A1+2` 的数字后面弹
    if (!freshArgument) return null;
    return { from, to: head.length, query, items: FORMULA_FUNCTIONS.slice(0, EMPTY_QUERY_LIMIT) };
  }

  // 纯数字查询只在「参数起始位置」提示整行引用，避免 `=A1+2` 时冒出 `2:2`
  if (/^\d+$/.test(query) && !argumentStart) return null;

  const upper = query.toUpperCase();
  const startsWith = FORMULA_FUNCTIONS.filter((fn) => fn.name.startsWith(upper));
  const contains = FORMULA_FUNCTIONS.filter(
    (fn) => !fn.name.startsWith(upper) && fn.name.includes(upper),
  );
  /**
   * 区域引用排在「前缀命中」之后、「包含命中」之前。
   * 用户在 `=SUM(B` 时想要的是整列 B，而不是名字里含 B 的 ABS/SUBTOTAL；
   * 反过来 `=SU` 时 SUM/SUMIF 这些前缀命中仍然排在最前面。
   */
  const references = referenceItems(query, usedRows, usedCols);
  const items = [...startsWith, ...references, ...contains].slice(0, QUERY_LIMIT);
  if (items.length === 0) return null;
  return { from, to: head.length, query, items };
}

/**
 * 光标前紧邻的引用记号（用于「拖选范围替换已有引用」）。
 *
 * 允许引用与光标之间夹着**右括号和空白**：`=SUM(B3:B4)` 时光标在末尾，
 * 若不允许 `)`，就会认不出前面已有引用而变成"追加"，实测表现为
 * `=SUM(B3:B4)` 拖选 E2:E5 后变成 `=SUM(B3:B4)E2:E5`（用户报的就是这个）。
 */
const REFERENCE_BEFORE_CARET =
  /(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(\)*[ \t]*)$/;

/**
 * 把一段引用写进公式文本（Excel 的"用鼠标选范围"行为）：
 * - 光标前（允许隔着右括号/空格）紧邻一个引用记号（`B3`、`B3:B4`、`B:B`、`2:2`）→ **替换**它；
 * - 否则在光标处**插入**。
 * 返回新文本与光标位置（停在引用之后，方便继续输入 `)`）。
 */
export function replaceOrInsertReference(
  text: string,
  caret: number,
  reference: string,
): { text: string; caret: number } {
  const position = Math.max(0, Math.min(caret, text.length));
  const head = text.slice(0, position);
  const match = REFERENCE_BEFORE_CARET.exec(head);
  if (match) {
    // match[2] 是引用与光标之间的右括号/空白：替换引用本身，保留它们
    const trailing = match[2] ?? "";
    const from = position - trailing.length - match[1].length;
    const next = text.slice(0, from) + reference + text.slice(from + match[1].length);
    return { text: next, caret: from + reference.length };
  }
  const next = text.slice(0, position) + reference + text.slice(position);
  return { text: next, caret: position + reference.length };
}

/**
 * 从公式里解析出引用的区域（用于在网格上高亮"当前公式的作用范围"）。
 *
 * 只处理当前工作表：带表名的引用（`Sheet1!A1`）仅在表名与 `sheetName` 相同时计入，
 * 跨表引用无法在当前表上画出范围，直接跳过。字符串字面量里的内容不参与解析。
 *
 * @param formula  公式文本（可带前导 `=`）
 * @param context  当前工作表名与已用行列数（整列 `B:B` / 整行 `2:2` 需要展开成具体范围）
 */
export function extractFormulaReferences(
  formula: string,
  context: { sheetName?: string; usedRows?: number; usedCols?: number } = {},
): Array<{ startRow: number; startCol: number; endRow: number; endCol: number }> {
  if (!formula || !formula.includes("=")) return [];
  const usedRows = Math.max(1, context.usedRows ?? 1);
  const usedCols = Math.max(1, context.usedCols ?? 1);
  // 去掉字符串字面量，避免把 "A1" 这种文本当成引用
  const cleaned = formula.replace(/"(?:[^"]|"")*"/g, '""');

  const ranges: Array<{ startRow: number; startCol: number; endRow: number; endCol: number }> = [];
  const cellRef = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
  const pattern = new RegExp(
    // 可选表名!，然后是「区域」或「单格」或「整列」或「整行」
    String.raw`(?:(?:'([^']+)'|([A-Za-z0-9_\u4e00-\u9fa5]+))!)?` +
      String.raw`(?:(${cellRef}):(${cellRef})|(${cellRef})|(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})|(\$?\d+):(\$?\d+))`,
    "g",
  );
  const toCell = (text: string): { row: number; col: number } | null => {
    const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(text);
    if (!match) return null;
    const col = columnIndex(match[1]);
    const row = Number(match[2]) - 1;
    if (col < 0 || row < 0) return null;
    return { row, col };
  };

  for (const match of cleaned.matchAll(pattern)) {
    const sheetName = match[1] ?? match[2];
    if (sheetName && context.sheetName && sheetName !== context.sheetName) continue;
    if (match[3] && match[4]) {
      const a = toCell(match[3]);
      const b = toCell(match[4]);
      if (!a || !b) continue;
      ranges.push({
        startRow: Math.min(a.row, b.row),
        startCol: Math.min(a.col, b.col),
        endRow: Math.max(a.row, b.row),
        endCol: Math.max(a.col, b.col),
      });
    } else if (match[5]) {
      const cell = toCell(match[5]);
      if (cell) ranges.push({ startRow: cell.row, startCol: cell.col, endRow: cell.row, endCol: cell.col });
    } else if (match[6] && match[7]) {
      const from = columnIndex(match[6].replace(/\$/g, ""));
      const to = columnIndex(match[7].replace(/\$/g, ""));
      if (from < 0 || to < 0) continue;
      ranges.push({
        startRow: 0,
        startCol: Math.min(from, to),
        endRow: usedRows - 1,
        endCol: Math.max(from, to),
      });
    } else if (match[8] && match[9]) {
      const from = Number(match[8].replace(/\$/g, "")) - 1;
      const to = Number(match[9].replace(/\$/g, "")) - 1;
      if (from < 0 || to < 0) continue;
      ranges.push({
        startRow: Math.min(from, to),
        startCol: 0,
        endRow: Math.max(from, to),
        endCol: usedCols - 1,
      });
    }
  }
  // 去重（同一个范围可能在公式里出现多次）
  const seen = new Set<string>();
  return ranges.filter((range) => {
    const key = `${range.startRow},${range.startCol},${range.endRow},${range.endCol}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/** 把选区（0 起、含端点）转成 Excel 引用文本。
 * - 单格 → `B2`
 * - 矩形 → `B2:D10`
 * - 整列（覆盖到已用区域最后一行且只有一列）→ `B:B`，贴近「对 B 列求和」的写法
 * - 整行（覆盖到已用区域最后一列且只有一行）→ `2:2`
 */
export function selectionReference(
  range: { startRow: number; startCol: number; endRow: number; endCol: number },
  context?: { usedRows?: number; usedCols?: number },
): string {
  const usedRows = context?.usedRows ?? 0;
  const usedCols = context?.usedCols ?? 0;
  const singleCol = range.startCol === range.endCol;
  const singleRow = range.startRow === range.endRow;
  if (singleCol && usedRows > 0 && range.startRow === 0 && range.endRow >= usedRows - 1) {
    const letter = columnName(range.startCol);
    return `${letter}:${letter}`;
  }
  if (singleRow && usedCols > 0 && range.startCol === 0 && range.endCol >= usedCols - 1) {
    return `${range.startRow + 1}:${range.startRow + 1}`;
  }
  const start = `${columnName(range.startCol)}${range.startRow + 1}`;
  const end = `${columnName(range.endCol)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/** 选中某个候选项后要写进输入框的文本，以及插入后光标应停在哪里 */
export function applyFormulaSuggestion(
  text: string,
  suggestion: FormulaSuggestion,
  fn: FormulaFunction,
): { text: string; caret: number } {
  // 引用类候选项（insert: "plain"）原样插入；函数插入 `NAME(`（无参函数 `NAME()`）
  const inserted =
    fn.insert === "plain" ? fn.name : fn.noArgs ? `${fn.name}()` : `${fn.name}(`;
  const next = text.slice(0, suggestion.from) + inserted + text.slice(suggestion.to);
  return { text: next, caret: suggestion.from + inserted.length };
}

/* ------------------------------------------------------------------ */
/* 区域引用提示：整列 / 整行 / 已用区域                                 */
/* ------------------------------------------------------------------ */

/** 0 起列号 → 列名（0 → A，26 → AA） */
export function columnName(index: number): string {
  let value = index + 1;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

/** 列名 → 0 起列号（A → 0；非法返回 -1） */
export function columnIndex(name: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(name)) return -1;
  let value = 0;
  for (const char of name.toUpperCase()) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  return value - 1;
}

/**
 * 当前查询词对应的区域引用候选。
 *
 * 目的是让用户能像 Excel 一样直接写整列/整行公式：
 * - 输入 `B` → `B:B`（整列 B）、`B2:B333`（B 列已用区域）
 * - 输入 `2` → `2:2`（整行 2）、`A2:C2`（第 2 行已用区域）
 * 数据范围由调用方给出（工作表已用行列数），因此不依赖公式引擎。
 */
function referenceItems(
  query: string,
  usedRows: number,
  usedCols: number,
): FormulaFunction[] {
  const upper = query.toUpperCase();
  const items: FormulaFunction[] = [];
  if (/^[A-Za-z]{1,3}$/.test(upper)) {
    const col = columnIndex(upper);
    if (col < 0) return items;
    const letter = columnName(col);
    items.push({
      name: `${letter}:${letter}`,
      signature: `${letter}:${letter}`,
      description: `整列 ${letter}（整列求和/统计）`,
      insert: "plain",
    });
    if (usedRows >= 2) {
      items.push({
        name: `${letter}2:${letter}${usedRows}`,
        signature: `${letter}2:${letter}${usedRows}`,
        description: `${letter} 列已用区域（第 2~${usedRows} 行）`,
        insert: "plain",
      });
    }
    return items;
  }
  if (/^\d{1,7}$/.test(query)) {
    const row = Number(query);
    if (row < 1) return items;
    items.push({
      name: `${row}:${row}`,
      signature: `${row}:${row}`,
      description: `整行 ${row}`,
      insert: "plain",
    });
    if (usedCols >= 2) {
      items.push({
        name: `A${row}:${columnName(usedCols - 1)}${row}`,
        signature: `A${row}:${columnName(usedCols - 1)}${row}`,
        description: `第 ${row} 行已用区域（A~${columnName(usedCols - 1)} 列）`,
        insert: "plain",
      });
    }
  }
  return items;
}
