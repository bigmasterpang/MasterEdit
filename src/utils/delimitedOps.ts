/**
 * CSV / TSV 的行列结构操作（插入 / 删除行、插入 / 删除列）。
 *
 * CSV 没有"结构"这一层，行列操作最终都体现为**文本改写**，因此这里按原文切片、
 * 尽量保留用户原有的书写风格：
 * - 换行符沿用文件本身的（CRLF / LF 混用也按行原样保留）；
 * - **按引号规则切分字段**，引号内的分隔符不当分隔符，重新拼接时保留原引号与内容；
 * - 不改动其它行的内容（不做规范化、不重排引号），只做插入或删除。
 *
 * 结构改动是整体文本替换，撤销由 `sheetHistory` 的内容步骤负责（一步还原）。
 */

/** 单行按分隔符切成原始字段（尊重双引号；返回原文，便于原样拼回） */
export function splitRawFields(line: string, delimiter: string): string[] {
  const fields: string[] = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const ch = line[index];
    if (quoted) {
      if (ch === '"') {
        if (line[index + 1] === '"') {
          current += '""';
          index += 1;
        } else {
          quoted = false;
          current += ch;
        }
      } else {
        current += ch;
      }
    } else if (ch === '"' && current.trim() === "") {
      quoted = true;
      current += ch;
    } else if (line.startsWith(delimiter, index)) {
      fields.push(current);
      current = "";
      index += delimiter.length - 1;
    } else {
      current += ch;
    }
  }
  fields.push(current);
  return fields;
}

/** 把文本切成「行 + 该行原本的换行符」，末行可能没有换行符 */
function splitLines(text: string): Array<{ body: string; eol: string }> {
  const lines: Array<{ body: string; eol: string }> = [];
  let start = 0;
  for (let index = 0; index < text.length; index += 1) {
    const ch = text[index];
    if (ch === "\n") {
      const hasCr = index > start && text[index - 1] === "\r";
      lines.push({ body: text.slice(start, hasCr ? index - 1 : index), eol: hasCr ? "\r\n" : "\n" });
      start = index + 1;
    }
  }
  if (start < text.length) lines.push({ body: text.slice(start), eol: "" });
  return lines;
}

function joinLines(lines: Array<{ body: string; eol: string }>): string {
  return lines.map((line) => line.body + line.eol).join("");
}

/** 在某行上/下插入一个空行 */
export function insertDelimitedRow(text: string, row: number, where: "above" | "below"): string {
  const lines = splitLines(text);
  const at = Math.max(0, Math.min(where === "above" ? row : row + 1, lines.length));
  const eol = lines[Math.max(0, at - 1)]?.eol || lines[0]?.eol || "\r\n";
  lines.splice(at, 0, { body: "", eol });
  return joinLines(lines);
}

/** 删除某一行（连同它的换行符） */
export function deleteDelimitedRow(text: string, row: number): string {
  const lines = splitLines(text);
  if (row < 0 || row >= lines.length) return text;
  const isLast = row === lines.length - 1;
  lines.splice(row, 1);
  // 删掉的是最后一行时，把新的末行换行符去掉，避免多出一个空行
  if (isLast && lines.length > 0) lines[lines.length - 1] = { ...lines[lines.length - 1], eol: "" };
  return joinLines(lines);
}

/** 在某一列左/右插入一个空列（每行对应位置插入一个空字段） */
export function insertDelimitedColumn(
  text: string,
  col: number,
  delimiter: string,
  where: "left" | "right",
): string {
  const lines = splitLines(text);
  const at = Math.max(0, where === "left" ? col : col + 1);
  for (const line of lines) {
    if (line.body === "" && lines.length > 1) continue; // 空行保持空行，不凭空造出分隔符
    const fields = splitRawFields(line.body, delimiter);
    fields.splice(Math.min(at, fields.length), 0, "");
    line.body = fields.join(delimiter);
  }
  return joinLines(lines);
}

/** 删除某一列（每行删掉对应字段；只剩一列时清空该行内容） */
export function deleteDelimitedColumn(text: string, col: number, delimiter: string): string {
  const lines = splitLines(text);
  for (const line of lines) {
    if (line.body === "") continue;
    const fields = splitRawFields(line.body, delimiter);
    if (col < 0 || col >= fields.length) continue;
    fields.splice(col, 1);
    line.body = fields.length === 0 ? "" : fields.join(delimiter);
  }
  return joinLines(lines);
}
