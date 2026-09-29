import { createDelimitedTable, serializeDelimitedRow } from "./delimited";

/**
 * 把表格里改动的单元格写回 CSV / TSV 文本：只替换被改的那一行，其余内容逐字节保持原样。
 *
 * 实现要点：
 * - 每次都从**当前文本**重新建立惰性索引，因此偏移永远与文本一致（可编辑的 CSV 有 10MB 上限，
 *   一次扫描 ≤25ms，用户感知不到）；
 * - 同一行的多个单元格编辑合并成一次替换，避免重复扫描；
 * - 行号越界（文件被外部改小）时跳过该行，绝不破坏内容。
 *
 * 返回 null 表示没有任何可应用的编辑（调用方保持原内容不变）。
 */
export function applyCellEditsToDelimited(
  content: string,
  edits: Array<{ row: number; col: number; value: string }>,
  options?: { delimiter?: string; ext?: string },
): string | null {
  if (edits.length === 0) return null;
  const table = createDelimitedTable(content, options);

  const byRow = new Map<number, Array<{ col: number; value: string }>>();
  for (const edit of edits) {
    const list = byRow.get(edit.row);
    if (list) list.push(edit);
    else byRow.set(edit.row, [edit]);
  }

  let text = content;
  let changed = false;
  // 从下往上替换：改动靠下的行只影响它后面的偏移，前面各行的区间始终有效
  for (const row of [...byRow.keys()].sort((a, b) => b - a)) {
    const range = table.rowRange(row);
    if (!range) continue;
    const cells = (table.rowAt(row) ?? []).map((cell) => cell.v);
    for (const edit of byRow.get(row) ?? []) {
      while (cells.length <= edit.col) cells.push("");
      cells[edit.col] = edit.value;
    }
    const line = serializeDelimitedRow(cells, table.delimiter);
    if (line === content.slice(range[0], range[1])) continue;
    text = text.slice(0, range[0]) + line + text.slice(range[1]);
    changed = true;
  }
  return changed ? text : null;
}
