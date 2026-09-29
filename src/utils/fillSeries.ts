/**
 * Excel 式「拖拽填充」的规律推断。
 *
 * 输入是**原选区的样本值**（按填充方向排列：向下填充就是从上到下），输出接下来要填的若干个值。
 * 支持这些规律（与 Excel 的默认行为一致，不需要按 Ctrl）：
 * 1. **纯数字等差数列**：`1,2,3` → `4,5,6`；`1,3,5` → `7,9`；只有一个数字时**原样复制**（Excel 默认也是复制）；
 * 2. **带尾号的文本**：`第1项` → `第2项`、`A01` → `A02`（保留前缀、后缀与补零位数）；
 * 3. **公式**：按填充方向**平移相对引用**（`=B2*C2` 向下填充 → `=B3*C3`），`$` 锁定的行列不动；
 * 4. 其它情况（纯文本、日期文本、无规律的混合）：**循环复制**样本。
 *
 * 只做"看得懂"的规律，不猜复杂序列（不做日期推算、不做自定义列表）—— 与项目"简单查看编辑"的定位一致。
 */

export type FillDirection = "down" | "up" | "right" | "left";

/** 数字识别：允许前后空白与千分位以外的普通写法 */
function parseNumber(text: string): number | null {
  const trimmed = text.trim();
  if (trimmed === "" || !/^[+-]?(\d+\.?\d*|\.\d+)$/.test(trimmed)) return null;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

/** 小数位数（用于结果格式化，避免 0.1+0.2 变成 0.30000000000000004） */
function decimalsOf(text: string): number {
  const match = /\.(\d+)/.exec(text.trim());
  return match ? match[1].length : 0;
}

/** 格式化数字：按样本的小数位数输出，去掉多余的 0 */
function formatNumber(value: number, decimals: number): string {
  const fixed = value.toFixed(Math.min(10, Math.max(0, decimals)));
  return decimals > 0 ? fixed.replace(/0+$/, "").replace(/\.$/, "") : fixed;
}

/** 带尾号文本的识别：前缀 + 数字 + 后缀 */
function parseTrailingNumber(text: string): { prefix: string; digits: string; suffix: string } | null {
  const match = /^(.*?)(\d+)(\D*)$/.exec(text);
  if (!match) return null;
  // 纯数字交给数字分支处理
  if (match[1] === "" && match[3] === "") return null;
  return { prefix: match[1], digits: match[2], suffix: match[3] };
}

/** 等差数列：返回步长；样本不足或不成等差时返回 null（null 表示"按复制处理"） */
function detectStep(values: number[]): number | null {
  if (values.length < 2) return null;
  const step = values[1] - values[0];
  for (let index = 2; index < values.length; index += 1) {
    if (Math.abs(values[index] - values[index - 1] - step) > 1e-9) return null;
  }
  return step;
}

/** 平移公式里的相对引用（`$` 锁定的行列不动；字符串字面量、函数名不动） */
export function shiftFormula(formula: string, rowDelta: number, colDelta: number): string {
  if (rowDelta === 0 && colDelta === 0) return formula;
  // 字符串字面量整体跳过（避免把 "A1" 这种文本当成引用平移）
  const parts = formula.split(/("(?:[^"]|"")*")/);
  return parts
    .map((part) => {
      if (part.startsWith('"')) return part;
      return part.replace(
        /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g,
        (
          whole: string,
          colLock: string,
          colText: string,
          rowLock: string,
          rowText: string,
          offset: number,
        ) => {
          const before = offset > 0 ? part[offset - 1] : "";
          const after = part[offset + whole.length] ?? "";
          // 紧邻字母/数字/下划线/点 → 是更长的标识符（LOG10、Sheet1）的一部分，不动
          if (/[A-Za-z0-9_.]/.test(before)) return whole;
          // 后面紧跟 `(` → 是函数名（LOG10(、SUM( 不会走到这里），不动
          if (after === "(") return whole;
          const col = colIndex(colText);
          const row = Number(rowText);
          const nextCol = colLock ? col : col + colDelta;
          const nextRow = rowLock ? row : row + rowDelta;
          if (nextCol < 1 || nextRow < 1) return whole;
          return `${colLock}${colName(nextCol)}${rowLock}${nextRow}`;
        },
      );
    })
    .join("");
}

function colIndex(text: string): number {
  let value = 0;
  for (const ch of text.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64);
  return value;
}

function colName(index: number): string {
  let value = index;
  let name = "";
  while (value > 0) {
    const rest = (value - 1) % 26;
    name = String.fromCharCode(65 + rest) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

/**
 * 根据样本推断接下来 `count` 个值。
 *
 * @param samples   原选区的样本值（按填充方向排列）
 * @param count     需要生成的数量
 * @param direction 填充方向（决定公式相对引用的平移量与"倒着填"的顺序）
 */
export function fillSeries(samples: string[], count: number, direction: FillDirection): string[] {
  if (count <= 0) return [];
  const usable = samples.length > 0 ? samples : [""];
  const backwards = direction === "up" || direction === "left";
  /**
   * 统一按"样本原顺序"（上→下 / 左→右）处理：
   * - 向前填（down/right）：从**最后一个**样本往后接；
   * - 向后填（up/left）：从**第一个**样本往前接，最后把结果倒过来，
   *   这样返回值始终是"从上到下 / 从左到右"的顺序，调用方直接按顺序写即可。
   */
  const anchor = backwards ? usable[0] : usable[usable.length - 1];
  const sign = backwards ? -1 : 1;
  const rowDelta = direction === "down" ? 1 : direction === "up" ? -1 : 0;
  const colDelta = direction === "right" ? 1 : direction === "left" ? -1 : 0;
  const finish = (values: string[]): string[] => (backwards ? [...values].reverse() : values);

  // 1) 公式：平移相对引用（每前进一格平移一次）
  if (usable.every((value) => value.trim().startsWith("="))) {
    const out: string[] = [];
    for (let index = 0; index < count; index += 1) {
      const rounds = Math.floor(index / usable.length) + 1;
      out.push(shiftFormula(anchor, rowDelta * rounds, colDelta * rounds));
    }
    return finish(out);
  }

  // 2) 纯数字等差
  const numbers = usable.map(parseNumber);
  if (numbers.every((value) => value !== null)) {
    const step = detectStep(numbers as number[]);
    const decimals = Math.max(...usable.map(decimalsOf));
    const base = parseNumber(anchor) as number;
    const out: string[] = [];
    for (let index = 0; index < count; index += 1) {
      // 只有一个样本（或不成等差）时按 Excel 默认行为复制
      const value = step === null ? base : base + step * (index + 1) * sign;
      out.push(formatNumber(value, decimals));
    }
    return finish(out);
  }

  // 3) 带尾号的文本（第1项 / A01 / 3号楼）：**单个样本也递增**（Excel 对这类文本就是这样）
  const trailing = usable.map(parseTrailingNumber);
  if (trailing.every((value) => value !== null)) {
    const items = trailing as Array<{ prefix: string; digits: string; suffix: string }>;
    const values = items.map((item) => Number(item.digits));
    const step = detectStep(values) ?? 1;
    const width = Math.max(...items.map((item) => item.digits.length));
    const anchorItem = backwards ? items[0] : items[items.length - 1];
    const anchorValue = backwards ? values[0] : values[values.length - 1];
    const out: string[] = [];
    let overflow = false;
    for (let index = 0; index < count; index += 1) {
      const value = anchorValue + step * (index + 1) * sign;
      // 倒退到负数就整体退回循环复制（避免出现 "-1项" 或半截结果）
      if (value < 0) {
        overflow = true;
        break;
      }
      out.push(`${anchorItem.prefix}${String(value).padStart(width, "0")}${anchorItem.suffix}`);
    }
    if (!overflow) return finish(out);
    const fallback: string[] = [];
    for (let index = 0; index < count; index += 1) fallback.push(usable[index % usable.length]);
    return finish(fallback);
  }

  // 4) 其它：循环复制
  const out: string[] = [];
  for (let index = 0; index < count; index += 1) out.push(usable[index % usable.length]);
  return finish(out);
}
