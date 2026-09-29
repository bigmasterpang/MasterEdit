/**
 * RFC 4180 分隔文本（CSV / TSV）解析工具。
 *
 * 解析器是手写的字符状态机：对文本只做「一遍」线性扫描，不调用 split，
 * 几 MB 的文件也能在毫秒级完成；单元格类型推断与显示格式化都在同一遍
 * 之后按需执行，方便插件式地被表格视图（SheetGrid）消费。
 *
 * 大文件（几十万行 / 十几 MB）用法：createDelimitedTable 先只扫描「行首偏移」，
 * 不为任何单元格建对象（额外内存只有一张 number[] 偏移表），行内容由 rowAt
 * 按需解析并缓存最近访问的行，配合虚拟滚动可以瞬间打开整表；
 * parseDelimited 复用同一套扫描 / 行解析逻辑，保证两条路径结果永远一致。
 */
import type { SheetCell, SheetCellType } from "../types";

/* ------------------------------------------------------------------ */
/* 常量                                                                */
/* ------------------------------------------------------------------ */

/** 默认最多解析的数据行数（仅 parseDelimited 使用；惰性行索引没有行数上限） */
const DEFAULT_MAX_ROWS = 20000;
/** 默认最多解析的列数（超出即截断） */
const DEFAULT_MAX_COLS = 500;
/** 嗅探分隔符时最多采样的非空行数 */
const SNIFF_LINES = 20;
/** 候选分隔符：数组顺序即「平局时的优先级」（逗号最优先） */
const CANDIDATE_DELIMITERS = [",", ";", "\t"] as const;
/** rowAt 的行缓存容量：虚拟滚动会反复取相邻行，缓存让同一行不必重复解析 */
const ROW_CACHE_SIZE = 256;

/* ------------------------------------------------------------------ */
/* 类型判定用的正则                                                     */
/* ------------------------------------------------------------------ */

const BOOL_RE = /^(?:true|false)$/i;

/** Excel / LibreOffice 的错误值：#N/A、#VALUE!、#REF!、#DIV/0! … */
const ERROR_RE =
  /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|BLOCKED!|CONNECT!|FIELD!|UNKNOWN!|BUSY!|SYNC!)$/i;

/**
 * 严格数字：整数 / 小数 / 「.5」/「1.」/ 科学计数法。
 * 这里比 RFC 语义略宽地允许前导 `+`，因为 formatCellForDisplay 明确要求
 * 「去掉前导 +」，若判定阶段不接受 `+`，那条规则就永远不会被触发。
 */
const NUMBER_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

/**
 * 千分位数字：分隔符后必须「恰好 3 位」且至少出现一组，
 * 这样 `1,234` / `12,345,678.9` 会被判为数字，而欧式小数写法
 * `1,23`（只有 2 位）不会，避免「1,23」这种歧义被误判。
 */
const GROUPED_NUMBER_RE = /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?(?:[eE][+-]?\d+)?$/;

/** ISO 风格日期：YYYY-MM-DD / YYYY/MM/DD，可选时间部分（含 Z / 时区偏移） */
const DATE_RE =
  /^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d{1,3})?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;

/* ------------------------------------------------------------------ */
/* 分隔符嗅探                                                          */
/* ------------------------------------------------------------------ */

/**
 * 统计一行中「引号之外」某个字符出现的次数。
 * 引号内的分隔符不算数；`""` 是转义，需要整体跳过，否则会误判引号的开闭。
 */
function countOutsideQuotes(line: string, ch: string): number {
  const target = ch.charCodeAt(0);
  let count = 0;
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const code = line.charCodeAt(i);
    if (code === 34 /* " */) {
      if (inQuotes && line.charCodeAt(i + 1) === 34) {
        i += 1; // "" 转义：跳过第二个引号，避免把引号状态翻回去
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && code === target) count += 1;
  }
  return count;
}

/**
 * 推断分隔符：`tsv` / `tab` 扩展名直接返回制表符；其余情况采样前 20 个
 * 非空行，用「引号外计数」得到每个候选分隔符切出的列数，
 * 取「最常见列数出现频率最高」的那个（一致性优先，列数多者次之），
 * 谁都切不出多列时兜底为逗号。
 */
export function detectDelimiter(text: string, ext: string): string {
  const kind = ext.replace(/^\./, "").trim().toLowerCase();
  if (kind === "tsv" || kind === "tab") return "\t";

  const counts: Record<string, number[]> = { ",": [], ";": [], "\t": [] };
  let sampled = 0;
  let lineStart = 0;
  let i = 0;
  const n = text.length;

  // 手工按行扫描（而不是 split），一旦采够样本立刻停止，避免为几 MB 文本分配行数组
  while (i <= n && sampled < SNIFF_LINES) {
    const code = i < n ? text.charCodeAt(i) : -1; // -1：文本结束，强制收尾最后一行
    if (code !== 10 /* \n */ && code !== 13 /* \r */ && code !== -1) {
      i += 1;
      continue;
    }
    const line = text.slice(lineStart, i);
    i += code === 13 && text.charCodeAt(i + 1) === 10 ? 2 : 1; // CRLF 视为一个换行
    lineStart = i;
    if (line.trim().length === 0) continue; // 空行不参与统计
    sampled += 1;
    for (const candidate of CANDIDATE_DELIMITERS) {
      counts[candidate].push(countOutsideQuotes(line, candidate) + 1);
    }
  }

  let best = ",";
  let bestScore = -1;
  for (const candidate of CANDIDATE_DELIMITERS) {
    const perLine = counts[candidate];
    if (perLine.length === 0) continue;
    const freq = new Map<number, number>();
    for (const value of perLine) freq.set(value, (freq.get(value) ?? 0) + 1);
    // 取出现次数最多的列数；次数相同则取列数更大的，避免选中「每行都只有 1 列」的候选
    let modeCols = 0;
    let modeFreq = 0;
    for (const [value, hits] of freq) {
      if (hits > modeFreq || (hits === modeFreq && value > modeCols)) {
        modeCols = value;
        modeFreq = hits;
      }
    }
    if (modeCols < 2) continue; // 切不出多列 → 不是这个分隔符
    const score = (modeFreq / perLine.length) * 1000 + Math.min(modeCols, 100);
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }
  return best;
}

/* ------------------------------------------------------------------ */
/* 单元格类型与显示文本                                                 */
/* ------------------------------------------------------------------ */

/** 推断单元格类型（仅用于判定，不改写原文） */
export function inferCellType(raw: string): SheetCellType {
  const text = raw.trim();
  if (text.length === 0) return "empty";
  if (BOOL_RE.test(text)) return "bool";
  if (ERROR_RE.test(text)) return "error";
  // 日期必须先于数字判断：2024-01-01 本身不满足数字正则，但顺序上更清晰
  if (DATE_RE.test(text)) return "date";
  if (NUMBER_RE.test(text) || GROUPED_NUMBER_RE.test(text)) return "number";
  return "text";
}

/**
 * 生成可直接显示的文本：
 * - 文本 / 日期 / 错误：原样保留（不做 trim，避免破坏用户原始数据的对齐信息）
 * - 数字：最小化归一 —— 只去掉前导 `+`，千分位与小数位数保持原样
 * - 布尔：统一成小写 true / false
 */
export function formatCellForDisplay(raw: string, type: SheetCellType): string {
  switch (type) {
    case "empty":
      return "";
    case "bool":
      return raw.trim().toLowerCase() === "true" ? "true" : "false";
    case "number": {
      const text = raw.trim();
      return text.charCodeAt(0) === 43 /* + */ ? text.slice(1) : text;
    }
    default:
      return raw;
  }
}

/* ------------------------------------------------------------------ */
/* 惰性行索引：扫描阶段（只记行首偏移，不建单元格）                       */
/* ------------------------------------------------------------------ */

/**
 * 归一化列上限：旧实现是「已收字段数 < maxCols 才收下」，等价于容量 ceil(maxCols)，
 * 所以小数上限（如 1.5 → 2 列）与非正数 / NaN（→ 0 列）都按同一规则折算成整数，
 * 保证 cols 永远是整数。
 */
function normalizeMaxCols(value: number): number {
  if (Number.isNaN(value)) return 0;
  return value > 0 ? Math.ceil(value) : 0;
}

/** 行索引扫描结果（模块内部使用，不对外暴露） */
interface RowIndexScan {
  /** 每行的起始下标（数组长度即数据行数） */
  starts: number[];
  /** 矩形宽度：最宽一行（受 maxCols 限制）的字段数 */
  cols: number;
  /** 是否有行的字段数超过 maxCols（那些行的尾部列被丢弃） */
  truncatedCols: boolean;
  /** 最后一行是否由「文本末尾」收尾（而非换行符）：精确判定 truncatedRows 时要用 */
  lastRowAtEof: boolean;
}

/**
 * 单遍字符状态机扫描：只记录每行的起始下标与字段数，**不为任何单元格分配对象**。
 *
 * 引号规则与 parseRowFields 逐字一致：
 * 1. `inQuotes` 是唯一的状态位；引号内的分隔符、`\n`、`\r` 都只是普通字符，
 *    所以字段内换行天然被支持，且跨物理行的字段仍属于同一行；
 * 2. `""` 在引号内表示一个字面量引号；
 * 3. 只有紧贴字段起始位置的引号才开启引号字段（RFC 语义）；
 * 4. 跳过文件头 BOM；结尾换行不多产生一行；末行没有换行符时正常收尾。
 *
 * 复杂度：O(n) 时间、O(行数) 空间（纯 number[] 偏移表，V8 下为紧凑 SMI 数组，
 * 每项 4～8 字节（取决于是否开启指针压缩）—— 12 万行约 0.5～1 MB，100 万行约 4～8 MB）。
 */
function scanRowIndex(text: string, delimCode: number, maxCols: number): RowIndexScan {
  const starts: number[] = [];
  const n = text.length;
  let cols = 0;
  let truncatedCols = false;

  // Excel 导出的 UTF-8 CSV 常带 BOM：跳过它，否则第一格会多一个不可见字符
  let i = n > 0 && text.charCodeAt(0) === 0xfeff ? 1 : 0;
  let rowStart = i; // 当前行的起始下标（遇到换行符后立刻更新）
  let fieldStart = i; // 当前字段起始下标：只有紧贴它的引号才算包裹引号
  let fieldCount = 0; // 当前行已完成的字段数
  let inQuotes = false; // 状态机核心：是否处于引号内部
  let quoted = false; // 当前字段是否为引号字段（收尾判定「是否停在行首」时要排除）

  /**
   * 收下一行：只记住起始下标与列数，行内容留到 rowAt 时再解析。
   * `truncatedCols` 不在这里判定，而是按「字段完成时」判（见下面两处 `fieldCount >= maxCols`），
   * 这样连「最终被丢弃的那一行的字段」也与旧实现的 pushField 行为一致。
   */
  const commitRow = (fields: number, start: number): void => {
    starts.push(start);
    const width = fields > maxCols ? maxCols : fields;
    if (width > cols) cols = width;
  };

  while (i < n) {
    const code = text.charCodeAt(i);

    if (inQuotes) {
      if (code === 34 /* " */) {
        if (text.charCodeAt(i + 1) === 34) {
          i += 2; // "" → 字段内的一个字面量引号
          continue;
        }
        inQuotes = false; // 收尾引号
        i += 1;
        continue;
      }
      i += 1; // 引号内的分隔符 / 换行 / 回车都只是内容
      continue;
    }

    if (code === 34 && i === fieldStart) {
      // 只有字段起始位置的引号才是包裹引号，其它位置的引号按普通字符处理
      quoted = true;
      inQuotes = true;
      i += 1;
      fieldStart = i;
      continue;
    }

    if (code === delimCode) {
      // 已收字段数达到列上限：本字段被丢弃（与旧实现 pushField 的判定同规则）
      if (fieldCount >= maxCols) truncatedCols = true;
      fieldCount += 1;
      i += 1;
      fieldStart = i;
      quoted = false;
      continue;
    }

    if (code === 10 /* \n */ || code === 13 /* \r */) {
      if (fieldCount >= maxCols) truncatedCols = true; // 行末字段同样可能被丢弃
      commitRow(fieldCount + 1, rowStart); // 换行符前的字段就是本行最后一个字段
      fieldCount = 0;
      i += code === 13 && text.charCodeAt(i + 1) === 10 ? 2 : 1; // CRLF 只算一个换行
      rowStart = i;
      fieldStart = i;
      quoted = false;
      continue;
    }

    i += 1;
  }

  // 收尾：文本正好停在行首说明末尾是换行符产生的空行，不再多算一行；
  // 否则补上最后一行（末行可能没有换行符）。与旧实现逐字一致：
  // 旧实现这里看的是「已收下的字段数」row.length，即字段数被 maxCols 截断后的值，
  // 所以 maxCols ≤ 0（字段全被丢弃）时它同样会把这最后一行当成空行跳过。
  const lastRowAtEof = !(i === fieldStart && Math.min(fieldCount, maxCols) === 0 && !quoted);
  if (lastRowAtEof) {
    if (fieldCount >= maxCols) truncatedCols = true; // 末行末字段被丢弃
    commitRow(fieldCount + 1, rowStart);
  }

  return { starts, cols, truncatedCols, lastRowAtEof };
}

/**
 * 解析某一行 [start, end) 的原始字段值（不做类型推断）。
 * - `end` 取「下一行起始下标」，最后一行取文本末尾；
 * - 引号规则与 scanRowIndex 逐字一致（含 `""` 转义、字段内换行、CRLF/CR/LF）；
 * - 最多返回 maxCols 个字段，超出的直接丢弃（列宽上限）。
 *
 * 因为区间由偏移表给出，这里不需要也不能重新全文扫描：行边界已经在扫描阶段定好，
 * 区间内不会出现「引号之外的换行」（否则扫描阶段会把它当作行边界）。
 */
function parseRowFields(
  text: string,
  start: number,
  end: number,
  delimCode: number,
  maxCols: number,
): string[] {
  const fields: string[] = [];
  let i = start;
  let fieldStart = start; // 当前字段内容的起始下标
  let fieldEnd = -1; // 引号字段的收尾引号下标（-1 表示未闭合/非引号字段）
  let quoted = false; // 当前字段是否由引号包裹（需要把 "" 还原为 "）
  let inQuotes = false; // 状态机核心：是否处于引号内部

  /** 收集当前字段（调用时 i 必须停在分隔符 / 换行符 / 文本末尾上） */
  const pushField = (): void => {
    const stop = fieldEnd >= 0 ? fieldEnd : i;
    if (fields.length < maxCols) {
      fields.push(quoted ? text.slice(fieldStart, stop).replace(/""/g, '"') : text.slice(fieldStart, stop));
    }
    // 超过 maxCols 的字段直接丢弃（与扫描阶段的列统计保持一致）
  };

  while (i < end) {
    const code = text.charCodeAt(i);

    if (inQuotes) {
      if (code === 34 /* " */) {
        if (text.charCodeAt(i + 1) === 34) {
          i += 2; // "" → 字段内的一个字面量引号
          continue;
        }
        inQuotes = false; // 收尾引号
        fieldEnd = i;
        i += 1;
        continue;
      }
      i += 1; // 引号内的分隔符 / 换行 / 回车都只是内容
      continue;
    }

    if (code === 34 && i === fieldStart) {
      quoted = true;
      inQuotes = true;
      i += 1;
      fieldStart = i;
      continue;
    }

    if (code === delimCode) {
      pushField();
      i += 1;
      fieldStart = i;
      fieldEnd = -1;
      quoted = false;
      continue;
    }

    if (code === 10 /* \n */ || code === 13 /* \r */) {
      pushField();
      return fields; // 行终止符：本行到此为止（最后一行末尾的换行也在此收束）
    }

    i += 1;
  }

  // 只有最后一行会走到这里（区间末尾即文本末尾）：停在行首说明本行没有内容
  const atLineStart = i === fieldStart && fields.length === 0 && !quoted && fieldEnd < 0;
  if (!atLineStart) pushField();
  return fields;
}

/* ------------------------------------------------------------------ */
/* 惰性行索引：对外接口                                                 */
/* ------------------------------------------------------------------ */

/** 惰性行索引：先单遍扫描只记录行首偏移，行内容按需解析（大文件友好） */
export interface DelimitedTable {
  /** 数据行数（不含结尾空行） */
  rows: number;
  /** 列数：按最宽行补齐/截断后的矩形宽度 */
  cols: number;
  /** 实际使用的分隔符 */
  delimiter: string;
  /** 是否因 maxCols 丢弃了某些行末尾的列 */
  truncatedCols: boolean;
  /** 按需取一行（0 起，已补齐/截断到 cols 列）；越界返回 undefined；内部缓存最近访问的行 */
  rowAt(index: number): SheetCell[] | undefined;
  /**
   * 某一行的字符区间 `[start, end)`（end 不含行尾换行符），用于「改一格 → 只替换这一行」的就地编辑；
   * 越界返回 undefined。
   */
  rowRange(index: number): [number, number] | undefined;
}

/**
 * 把一行字段序列化为 RFC 4180 文本：含分隔符 / 引号 / 换行 / 首尾空格时加引号，内部引号翻倍。
 * 表格内编辑 CSV 时用它重建被修改的那一行。
 */
export function serializeDelimitedRow(cells: string[], delimiter: string): string {
  return cells
    .map((value) => {
      const needsQuote =
        value.includes(delimiter) ||
        value.includes('"') ||
        value.includes("\n") ||
        value.includes("\r") ||
        value !== value.trim();
      return needsQuote ? `"${value.replace(/"/g, '""')}"` : value;
    })
    .join(delimiter);
}

/** 内部构造结果：把偏移表一并交给 parseDelimited，避免二次扫描 */
interface BuiltDelimitedTable {
  table: DelimitedTable;
  /** 行首偏移表（长度 = table.rows） */
  starts: number[];
  /** 最后一行是否由文本末尾收尾（而非换行符） */
  lastRowAtEof: boolean;
}

/**
 * 构造惰性行索引：createDelimitedTable 与 parseDelimited 共用的唯一入口，
 * 因此两条路径的扫描、补齐、类型推断规则完全一致（不存在第二套状态机）。
 *
 * rowAt 返回的数组是缓存实例（同一行重复取返回同一个对象），调用方只读使用。
 */
function buildDelimitedTable(
  text: string,
  options?: { delimiter?: string; ext?: string; maxCols?: number },
): BuiltDelimitedTable {
  const maxCols = normalizeMaxCols(options?.maxCols ?? DEFAULT_MAX_COLS);
  // 未显式给出分隔符时按内容嗅探（与 detectDelimiter 的兜底行为一致）
  const requested = options?.delimiter;
  const delimiter =
    requested && requested.length > 0 ? requested[0] : detectDelimiter(text, options?.ext ?? "");
  const delimCode = delimiter.charCodeAt(0);

  const scan = scanRowIndex(text, delimCode, maxCols);
  const starts = scan.starts;
  const rows = starts.length;
  const cols = scan.cols;
  /** 最近访问的行的解析结果：Map 的插入序即淘汰顺序，命中后重新插入实现 LRU */
  const cache = new Map<number, SheetCell[]>();

  const rowAt = (index: number): SheetCell[] | undefined => {
    if (!Number.isInteger(index) || index < 0 || index >= rows) return undefined;
    const hit = cache.get(index);
    if (hit !== undefined) {
      cache.delete(index);
      cache.set(index, hit);
      return hit;
    }
    // 行区间 = [本行起始, 下一行起始)，末行到文本末尾：全程查偏移表，不重新全文扫描
    const end = index + 1 < rows ? starts[index + 1] : text.length;
    const fields = parseRowFields(text, starts[index], end, delimCode, maxCols);
    // 矩形化：不足 cols 列用空单元格补齐（raw 为 ""，类型自然推断为 empty）
    const cells: SheetCell[] = new Array<SheetCell>(cols);
    for (let c = 0; c < cols; c += 1) {
      const raw = c < fields.length ? fields[c] : "";
      const type = inferCellType(raw);
      cells[c] = { v: formatCellForDisplay(raw, type), t: type };
    }
    cache.set(index, cells);
    if (cache.size > ROW_CACHE_SIZE) {
      const oldest = cache.keys().next();
      if (!oldest.done) cache.delete(oldest.value);
    }
    return cells;
  };

  /**
   * 某一行的字符区间 [start, end)，end **不含行尾换行符**（CR/LF/CRLF 都剥掉），
   * 便于「改一格 → 只替换这一行」的就地编辑。越界返回 undefined。
   */
  const rowRange = (index: number): [number, number] | undefined => {
    if (!Number.isInteger(index) || index < 0 || index >= rows) return undefined;
    const start = starts[index];
    let end = index + 1 < rows ? starts[index + 1] : text.length;
    while (end > start && (text.charCodeAt(end - 1) === 10 || text.charCodeAt(end - 1) === 13)) end -= 1;
    return [start, end];
  };

  return {
    table: { rows, cols, delimiter, truncatedCols: scan.truncatedCols, rowAt, rowRange },
    starts,
    lastRowAtEof: scan.lastRowAtEof,
  };
}

/**
 * 建立惰性行索引：单遍扫描只记录行首偏移与列数，行内容按需解析，没有行数上限。
 * 适合几十万行的大 CSV —— 扫描完即可拿到 rows / cols 交给虚拟滚动。
 */
export function createDelimitedTable(
  text: string,
  options?: { delimiter?: string; ext?: string; maxCols?: number },
): DelimitedTable {
  return buildDelimitedTable(text, options).table;
}

/* ------------------------------------------------------------------ */
/* 解析结果                                                            */
/* ------------------------------------------------------------------ */

export interface DelimitedParseResult {
  /** 矩形化的行数据（每行长度相同，等于最宽行的列数） */
  rows: SheetCell[][];
  /** 实际使用的分隔符 */
  delimiter: string;
  /** 是否因为 maxRows 而丢掉了后面的行 */
  truncatedRows: boolean;
  /** 是否因为 maxCols 而丢掉了某些行末尾的列 */
  truncatedCols: boolean;
}

/** 判断 from 之后是否还有非空白字符：用于确认截断标志是否名副其实 */
function hasContent(text: string, from: number): boolean {
  for (let i = from; i < text.length; i++) {
    if (text.charCodeAt(i) > 32) return true;
  }
  return false;
}

/**
 * 判定 truncatedRows：与改造前的语义逐字对应 —— 只有「被 maxRows 挡在外面的那一行
 * 真的有内容（或者它后面还有非空白内容）」才算截断，纯结尾空行、纯空字段行都不算。
 * `keptRows` 既是保留的行数，也是第一个被丢弃的行下标。
 */
function isTruncatedRows(
  text: string,
  starts: number[],
  rows: number,
  keptRows: number,
  lastRowAtEof: boolean,
  delimCode: number,
  maxCols: number,
): boolean {
  if (keptRows >= rows) return false; // 没有任何行被挡在外面
  const hasNext = keptRows + 1 < rows;
  if (!hasNext && lastRowAtEof) {
    // 被挡住的正是「由文本末尾收尾」的末行：旧实现在收尾分支里直接判为截断
    return true;
  }
  const end = hasNext ? starts[keptRows + 1] : text.length;
  const fields = parseRowFields(text, starts[keptRows], end, delimCode, maxCols);
  for (const field of fields) {
    if (field.length > 0) return true; // 该行有内容 → 确实丢了数据
  }
  // 该行自身是空的：再看它后面还有没有非空白内容（hasNext 为假时后面必然没有内容）
  return hasNext ? hasContent(text, end) : false;
}

/**
 * RFC 4180 解析：基于惰性行索引（createDelimitedTable）实现，取前 maxRows 行实例化。
 *
 * 1. 扫描、字段解析与 createDelimitedTable 完全同一套代码，所以
 *    `parseDelimited(text, { maxRows: n })` 的前 n 行与 `createDelimitedTable(text).rowAt(0..n-1)`
 *    逐格相同（列宽同取最宽行，受 maxCols 限制）；
 * 2. 达到 maxRows 后不再为后续行建单元格对象，但扫描仍是一次线性；
 * 3. 结尾换行不会多产生一行空行；末行没有换行符时会正常收尾；
 * 4. 行长度以「最宽的一行」为准补齐或截断，保证网格是矩形，方便虚拟滚动。
 *
 * 与改造前唯一有意保留的差别（也是「两条路径永远一致」的前提）：列宽与
 * truncatedCols 来自**整张表**（table.cols 统计全部行），而旧实现只统计被保留的
 * 前 maxRows 行。因此当「被丢弃的行」比被保留的行更宽时，新实现会多补出几列空
 * 单元格、truncatedCols 也可能由 false 变 true；行数、既有单元格内容与
 * truncatedRows 的判定则与旧实现逐字一致（maxCols=0 这类退化输入也一样）。
 */
export function parseDelimited(
  text: string,
  options?: { delimiter?: string; maxRows?: number; maxCols?: number },
): DelimitedParseResult {
  const maxCols = normalizeMaxCols(options?.maxCols ?? DEFAULT_MAX_COLS);
  // NaN 与旧实现一致地视为「不限行数」（旧实现的 >= 比较在 NaN 下恒为 false）
  const rawLimit = options?.maxRows ?? DEFAULT_MAX_ROWS;
  const limit = Number.isNaN(rawLimit) ? Infinity : rawLimit;

  const { table, starts, lastRowAtEof } = buildDelimitedTable(text, {
    delimiter: options?.delimiter,
    maxCols,
  });

  // 旧实现是「收满 maxRows 行后遇到下一行才停」，小数上限等价于向上取整（如 1.5 → 2 行）
  const rowCount = limit > 0 ? Math.min(table.rows, Math.ceil(limit)) : 0;
  const rows: SheetCell[][] = new Array<SheetCell[]>(rowCount);
  for (let r = 0; r < rowCount; r += 1) {
    rows[r] = table.rowAt(r) ?? [];
  }

  return {
    rows,
    delimiter: table.delimiter,
    truncatedRows: isTruncatedRows(
      text,
      starts,
      table.rows,
      rowCount,
      lastRowAtEof,
      table.delimiter.charCodeAt(0),
      maxCols,
    ),
    truncatedCols: table.truncatedCols,
  };
}
