import type { ChangeSpec, EditorState, Line } from "@codemirror/state";
import { EditorSelection } from "@codemirror/state";
import { EditorView } from "@codemirror/view";
import { getEditorView, withEditorView } from "./editorBridge";
import { useAppStore } from "../stores/appStore";
import { parseHeadings } from "./markdown";

/**
 * Markdown 编辑命令集合。
 * 所有命令都作用于当前光标 / 选中行，并尽量保持选区与撤销历史。
 */

export type ListKind = "bullet" | "ordered" | "task";
export type CalloutType =
  | "note"
  | "tip"
  | "important"
  | "warning"
  | "caution"
  | "info"
  | "success"
  | "question"
  | "quote"
  | "danger";

const HEADING_RE = /^(#{1,6})\s+/;
const TASK_RE = /^(\s*)([-*+])\s+\[([ xX])\]\s+/;
const BULLET_RE = /^(\s*)([-*+])\s+(?![[])/;
const ORDERED_RE = /^(\s*)(\d+)[.)]\s+/;
const QUOTE_RE = /^(\s*)>\s?/;

/* ------------------------------ 基础工具 ------------------------------ */

/** 取得选区覆盖的所有行（至少一行） */
function selectedLines(state: EditorState): Line[] {
  const sel = safeSelection(state);
  const start = state.doc.lineAt(sel.from);
  const endLine = state.doc.lineAt(sel.to);
  const lines: Line[] = [];
  for (let n = start.number; n <= endLine.number; n++) lines.push(state.doc.line(n));
  return lines;
}

/** 选区文本（用于替换） */
function selectionRange(state: EditorState): { from: number; to: number } {
  return safeSelection(state);
}

/**
 * 取安全的变更区间：始终 from <= to 且落在文档范围内。
 * 极端情况下（选区与文档不同步）CodeMirror 会因 from > to 抛 RangeError，
 * 统一在这里兜底，保证任何命令都不会因此中断。
 */
function clampChange(
  state: EditorState,
  from: number,
  to: number,
): { from: number; to: number } {
  const length = state.doc.length;
  const low = Math.max(0, Math.min(from, to, length));
  const high = Math.max(low, Math.min(Math.max(from, to), length));
  return { from: low, to: high };
}

/** 规范化当前选区 */
function safeSelection(state: EditorState): { from: number; to: number } {
  const sel = state.selection.main;
  return clampChange(state, sel.from, sel.to);
}

/** 批量修正变更列表中的越界 / 反向区间 */
function clampChanges(state: EditorState, changes: ChangeSpec[]): ChangeSpec[] {
  return changes.map((spec) => {
    if (!spec || typeof spec !== "object" || !("from" in spec)) return spec;
    const item = spec as { from?: number; to?: number };
    const from = typeof item.from === "number" ? item.from : 0;
    const to = typeof item.to === "number" ? item.to : from;
    return { ...item, ...clampChange(state, from, to) } as ChangeSpec;
  });
}

/** 单个变更的修正（返回 { from, to } + 原始 insert 字段） */
function safeChange(
  state: EditorState,
  from: number,
  to: number,
  insert: string,
): ChangeSpec {
  return { ...clampChange(state, from, to), insert };
}

/** 去掉行首的列表 / 引用标记，返回缩进 + 内容 */
function stripMarkers(text: string): string {
  return text
    .replace(/^(\s*)([-*+])\s+\[[ xX]\]\s+/, "$1")
    .replace(/^(\s*)([-*+])\s+/, "$1")
    .replace(/^(\s*)(\d+)[.)]\s+/, "$1")
    .replace(/^(\s*)>\s?/, "$1");
}

function isTaskLine(text: string): boolean {
  return TASK_RE.test(text);
}
function isBulletLine(text: string): boolean {
  return BULLET_RE.test(text) && !isTaskLine(text);
}
function isOrderedLine(text: string): boolean {
  return ORDERED_RE.test(text);
}
function isQuoteLine(text: string): boolean {
  return QUOTE_RE.test(text);
}

function applyChanges(view: EditorView, changes: ChangeSpec[]): void {
  if (changes.length === 0) return;
  view.dispatch({ changes: clampChanges(view.state, changes) });
}

/* ------------------------------ 内联格式 ------------------------------ */

/**
 * 切换包裹类格式（粗体、斜体、删除线、行内代码、上下标等）。
 * 已包裹则移除定界符，未包裹则包裹 —— 支持三种情况：
 * 1. 选区包含定界符：选中 `*text*` → 移除；
 * 2. 选区外侧紧邻定界符：光标/选区在 `*text*` 内 → 移除；
 * 3. 其它情况 → 包裹（空选区时扩展到所在单词）。
 */
export function toggleWrap(
  marker: string,
  placeholder = "",
  options: { word?: boolean } = {},
): void {
  withEditorView((view) => {
    const state = view.state;
    const length = marker.length;
    let { from, to } = safeSelection(state);

    // 空光标：扩展到所在单词（不含空白与定界符字符）
    if (from === to && options.word !== false) {
      const line = state.doc.lineAt(from);
      const isWordChar = (ch: string) =>
        ch.length > 0 && !/\s/.test(ch) && !marker.includes(ch);
      let start = from;
      let end = to;
      while (start > line.from && isWordChar(state.sliceDoc(start - 1, start))) start -= 1;
      while (end < line.to && isWordChar(state.sliceDoc(end, end + 1))) end += 1;
      if (end > start) {
        from = start;
        to = end;
      }
    }

    const selected = state.sliceDoc(from, to);

    // 情况 1：选区自带定界符
    if (
      selected.length >= length * 2 &&
      selected.startsWith(marker) &&
      selected.endsWith(marker)
    ) {
      const inner = selected.slice(length, selected.length - length);
      view.dispatch({
        changes: safeChange(state, from, to, inner),
        selection: { anchor: from, head: from + inner.length },
        scrollIntoView: true,
      });
      view.focus();
      return;
    }

    // 情况 2：选区外侧紧邻定界符
    const before = state.sliceDoc(Math.max(0, from - length), from);
    const after = state.sliceDoc(to, Math.min(state.doc.length, to + length));
    if (before === marker && after === marker) {
      const outerFrom = from - length;
      view.dispatch({
        changes: safeChange(state, outerFrom, to + length, selected),
        selection: { anchor: outerFrom, head: outerFrom + selected.length },
        scrollIntoView: true,
      });
      view.focus();
      return;
    }

    // 情况 3：包裹
    if (!selected && !placeholder) return;
    const text = selected || placeholder;
    view.dispatch({
      changes: safeChange(state, from, to, `${marker}${text}${marker}`),
      selection: {
        anchor: from + length,
        head: from + length + text.length,
      },
      scrollIntoView: true,
    });
    view.focus();
  });
}

/** 用定界符包裹选中文本；未选中时插入占位文本并选中（不做切换检测） */
export function wrapSelection(before: string, after = before, placeholder = ""): void {
  withEditorView((view) => {
    const { state } = view;
    const range = safeSelection(state);
    const selected = state.sliceDoc(range.from, range.to);
    const text = selected || placeholder;
    const insert = `${before}${text}${after}`;
    view.dispatch({
      changes: safeChange(state, range.from, range.to, insert),
      selection: {
        anchor: range.from + before.length,
        head: range.from + before.length + text.length,
      },
      scrollIntoView: true,
    });
    view.focus();
  });
}

/** 插入文本（替换选区） */
export function insertText(text: string): void {
  withEditorView((view) => {
    const state = view.state;
    const range = safeSelection(state);
    view.dispatch({
      changes: safeChange(state, range.from, range.to, text),
      selection: { anchor: range.from + text.length },
      scrollIntoView: true,
    });
    view.focus();
  });
}

/** 在光标处插入块级内容，必要时补空行 */
export function insertBlock(text: string): void {
  withEditorView((view) => {
    const { state } = view;
    const range = safeSelection(state);
    const lineStart = state.doc.lineAt(range.from).from;
    const before = state.sliceDoc(Math.max(0, lineStart - 2), lineStart);
    const prefix = lineStart === 0 || before.endsWith("\n\n") ? "" : "\n";
    const insert = `${prefix}${text}\n`;
    view.dispatch({
      changes: safeChange(state, range.from, range.to, insert),
      selection: { anchor: range.from + insert.length },
      scrollIntoView: true,
    });
    view.focus();
  });
}

export function toggleBold(): void {
  toggleWrap("**", "粗体");
}
export function toggleItalic(): void {
  toggleWrap("*", "斜体");
}
export function toggleStrikethrough(): void {
  toggleWrap("~~", "删除线");
}
export function toggleInlineCode(): void {
  toggleWrap("`", "代码");
}
export function toggleSuperscript(): void {
  toggleWrap("^", "上标");
}
export function toggleSubscript(): void {
  toggleWrap("~", "下标");
}
/** 高亮 ==text== */
export function toggleHighlight(): void {
  toggleWrap("==", "高亮");
}
/** 下划线 ++text++ */
export function toggleUnderline(): void {
  toggleWrap("++", "下划线");
}

/* ------------------------------ 块级结构 ------------------------------ */

/** 设置标题级别；level=0 表示普通段落；再次点击同级标题可取消 */
export function setHeading(level: number): void {
  withEditorView((view) => {
    const { state } = view;
    const changes: ChangeSpec[] = [];
    for (const line of selectedLines(state)) {
      const text = line.text;
      const existing = HEADING_RE.exec(text);
      let next: string;
      if (existing && (existing[1].length === level || level === 0)) {
        next = text.slice(existing[0].length);
      } else if (level === 0) {
        next = text;
      } else {
        next = `${"#".repeat(level)} ${existing ? text.slice(existing[0].length) : text}`;
      }
      if (next !== text) changes.push({ from: line.from, to: line.to, insert: next });
    }
    applyChanges(view, changes);
    view.focus();
  });
}

/**
 * 整体提升 / 降低标题级别。
 * 级别 0 表示普通段落：
 * - 提升：段落 → H6，H6 → H5 … H2 → H1（H1 已是最高，保持不变）
 * - 降低：H1 → H2 … H6 → 普通段落（段落无法再降）
 */
export function shiftHeading(delta: number): void {
  withEditorView((view) => {
    const { state } = view;
    const changes: ChangeSpec[] = [];
    for (const line of selectedLines(state)) {
      const text = line.text;
      const existing = HEADING_RE.exec(text);
      const current = existing ? existing[1].length : 0;
      const body = existing ? text.slice(existing[0].length) : text;

      let level: number;
      if (delta < 0) {
        level = current === 0 ? 6 : Math.max(1, current - 1);
      } else {
        if (current === 0) continue; // 段落无法再降级
        level = current >= 6 ? 0 : current + 1;
      }

      const next = level === 0 ? body : `${"#".repeat(level)} ${body}`;
      if (next !== text) changes.push({ from: line.from, to: line.to, insert: next });
    }
    applyChanges(view, changes);
    view.focus();
  });
}

/** 切换列表类型（再次应用同类型则取消列表） */
export function toggleList(kind: ListKind): void {
  withEditorView((view) => {
    const { state } = view;
    const lines = selectedLines(state);
    const allSame = lines.every((line) => {
      const text = line.text;
      if (kind === "task") return isTaskLine(text);
      if (kind === "bullet") return isBulletLine(text);
      return isOrderedLine(text);
    });

    const changes: ChangeSpec[] = [];
    let index = 0;
    for (const line of lines) {
      const text = line.text;
      const stripped = stripMarkers(text);
      const indent = /^\s*/.exec(stripped)?.[0] ?? "";
      const body = stripped.slice(indent.length);
      let next: string;
      if (allSame) {
        next = stripped;
      } else if (kind === "task") {
        next = `${indent}- [ ] ${body}`;
      } else if (kind === "bullet") {
        next = `${indent}- ${body}`;
      } else {
        next = `${indent}${index + 1}. ${body}`;
      }
      index += 1;
      if (next !== text) changes.push({ from: line.from, to: line.to, insert: next });
    }
    applyChanges(view, changes);
    view.focus();
  });
}

/** 切换引用块 */
export function toggleQuote(): void {
  withEditorView((view) => {
    const { state } = view;
    const lines = selectedLines(state);
    const allQuoted = lines.every((line) => isQuoteLine(line.text));
    const changes: ChangeSpec[] = lines.map((line) => {
      const text = line.text;
      const next = allQuoted
        ? text.replace(/^(\s*)>\s?/, "$1")
        : `> ${text}`;
      return { from: line.from, to: line.to, insert: next };
    });
    applyChanges(view, changes);
    view.focus();
  });
}

/** 插入提示块：> [!NOTE] 标题 */
export function insertCallout(type: CalloutType = "note"): void {
  withEditorView((view) => {
    const { state } = view;
    const range = safeSelection(state);
    const selected = state.sliceDoc(range.from, range.to);
    const label = type.toUpperCase();

    if (!selected) {
      const template = `> [!${label}]\n> 在此输入内容\n`;
      view.dispatch({
        changes: safeChange(state, range.from, range.to, template),
        selection: {
          anchor: range.from + `> [!${label}]\n> `.length,
          head: range.from + `> [!${label}]\n> `.length + 6,
        },
        scrollIntoView: true,
      });
      view.focus();
      return;
    }

    const lines = selected.split(/\r?\n/);
    const body =
      lines.length === 1
        ? `> [!${label}] ${lines[0]}`
        : `> [!${label}]\n${lines.map((l) => `> ${l}`).join("\n")}`;
    view.dispatch({
      changes: safeChange(state, range.from, range.to, body),
      selection: { anchor: range.from + body.length },
      scrollIntoView: true,
    });
    view.focus();
  });
}

/** 插入表格 */
export function insertTable(cols = 3, rows = 3): void {
  const width = Math.max(1, Math.min(8, cols));
  const height = Math.max(1, Math.min(20, rows));
  const header = `| ${Array.from({ length: width }, (_, i) => `列 ${i + 1}`).join(" | ")} |`;
  const divider = `| ${Array.from({ length: width }, () => "---").join(" | ")} |`;
  const body = Array.from(
    { length: height },
    () => `| ${Array.from({ length: width }, () => "   ").join(" | ")} |`,
  ).join("\n");
  insertBlock(`${header}\n${divider}\n${body}`);
}

/** 插入水平分割线 */
export function insertHorizontalRule(): void {
  insertBlock("---");
}

/** 插入代码块（围绕选中内容） */
export function insertCodeBlock(lang = ""): void {
  withEditorView((view) => {
    const { state } = view;
    const range = safeSelection(state);
    const selected = state.sliceDoc(range.from, range.to) || "代码";
    const fence = "```";
    const block = `${fence}${lang}\n${selected}\n${fence}`;
    view.dispatch({
      changes: safeChange(state, range.from, range.to, block),
      selection: {
        anchor: range.from + fence.length + lang.length + 1,
        head: range.from + fence.length + lang.length + 1 + selected.length,
      },
      scrollIntoView: true,
    });
    view.focus();
  });
}

/** 插入当前日期时间 */
export function insertDateTime(): void {
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  insertText(
    `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(
      now.getHours(),
    )}:${pad(now.getMinutes())}`,
  );
}

/** 依据文档标题生成目录 */
export function insertToc(): void {
  withEditorView((view) => {
    const headings = parseHeadings(view.state.doc.toString());
    if (headings.length === 0) {
      insertText("- （当前文档没有标题）");
      return;
    }
    const lines = headings
      .map((h) => {
        const indent = "  ".repeat(Math.max(0, h.level - 1));
        return `${indent}- [${h.text || "未命名"}](#${h.id})`;
      })
      .join("\n");
    insertBlock(`## 目录\n\n${lines}`);
  });
}

/* ------------------------------ 行操作 ------------------------------ */

/** 上下移动选中行 */
export function moveLines(direction: -1 | 1): void {
  withEditorView((view) => {
    const { state } = view;
    const sel = safeSelection(state);
    const startLine = state.doc.lineAt(sel.from);
    const endLine = state.doc.lineAt(sel.to);

    if (direction === -1) {
      if (startLine.number === 1) return;
      const prev = state.doc.line(startLine.number - 1);
      const block = state.sliceDoc(startLine.from, endLine.to);
      const shift = prev.length + 1;
      view.dispatch({
        changes: safeChange(state, prev.from, endLine.to, `${block}\n${prev.text}`),
        selection: EditorSelection.range(
          Math.max(0, sel.from - shift),
          Math.max(0, sel.to - shift),
        ),
        scrollIntoView: true,
      });
      return;
    }

    if (endLine.number === state.doc.lines) return;
    const next = state.doc.line(endLine.number + 1);
    const block = state.sliceDoc(startLine.from, endLine.to);
    const shift = next.length + 1;
    view.dispatch({
      changes: safeChange(state, startLine.from, next.to, `${next.text}\n${block}`),
      selection: EditorSelection.range(sel.from + shift, sel.to + shift),
      scrollIntoView: true,
    });
  });
}

/** 复制选中行 */
export function duplicateLines(): void {
  withEditorView((view) => {
    const { state } = view;
    const sel = safeSelection(state);
    const startLine = state.doc.lineAt(sel.from);
    const endLine = state.doc.lineAt(sel.to);
    const block = state.sliceDoc(startLine.from, endLine.to);
    const insert = `\n${block}`;
    view.dispatch({
      changes: safeChange(state, endLine.to, endLine.to, insert),
      selection: EditorSelection.range(sel.from, sel.to),
      scrollIntoView: true,
    });
  });
}

/** 删除选中行 */
export function deleteLines(): void {
  withEditorView((view) => {
    const { state } = view;
    const sel = safeSelection(state);
    const startLine = state.doc.lineAt(sel.from);
    const endLine = state.doc.lineAt(sel.to);
    let from = startLine.from;
    let to = endLine.to;
    if (to < state.doc.length) to += 1;
    else if (from > 0) from -= 1;
    view.dispatch({
      changes: safeChange(state, from, to, ""),
      selection: { anchor: Math.min(from, state.doc.length) },
      scrollIntoView: true,
    });
    view.focus();
  });
}


/** 大小写转换 */
export function transformCase(kind: "upper" | "lower" | "title"): void {
  withEditorView((view) => {
    const { state } = view;
    const range = selectionRange(state);
    const selected = state.sliceDoc(range.from, range.to);
    if (!selected) return;
    let next = selected;
    if (kind === "upper") next = selected.toUpperCase();
    else if (kind === "lower") next = selected.toLowerCase();
    else
      next = selected.replace(/\b([a-z])(\w*)/gi, (_, first: string, rest: string) =>
        `${first.toUpperCase()}${rest}`,
      );
    view.dispatch({
      changes: safeChange(state, range.from, range.to, next),
      selection: { anchor: range.from, head: range.from + next.length },
    });
    view.focus();
  });
}

/* ------------------------------ 视图辅助 ------------------------------ */


/** 跳转到指定行（0 起） */
export function scrollToLine(line: number): void {
  withEditorView((view) => {
    const doc = view.state.doc;
    const clamped = Math.min(Math.max(1, line + 1), doc.lines);
    const info = doc.line(clamped);
    view.dispatch({
      selection: { anchor: info.from },
      effects: EditorView.scrollIntoView(info.from, { y: "start", yMargin: 24 }),
    });
    view.focus();
  });
}

/** 插入数学公式：选中内容作为公式体，未选中则插入模板 */
export function insertMath(display: boolean): void {
  withEditorView((view) => {
    const { state } = view;
    const range = safeSelection(state);
    const selected = state.sliceDoc(range.from, range.to).trim();
    if (display) {
      const body = selected || "\\int_{0}^{\\infty} e^{-x^2}\\,dx = \\frac{\\sqrt{\\pi}}{2}";
      const block = `$$\n${body}\n$$`;
      view.dispatch({
        changes: safeChange(state, range.from, range.to, block),
        selection: {
          anchor: range.from + 3,
          head: range.from + 3 + body.length,
        },
        scrollIntoView: true,
      });
    } else {
      const body = selected || "E = mc^2";
      const inline = `$${body}$`;
      view.dispatch({
        changes: safeChange(state, range.from, range.to, inline),
        selection: {
          anchor: range.from + 1,
          head: range.from + 1 + body.length,
        },
        scrollIntoView: true,
      });
    }
    view.focus();
  });
}

/** 插入 Mermaid 图表模板 */
export function insertMermaid(): void {
  const template = [
    "```mermaid",
    "graph TD",
    "    A[开始] --> B{条件判断}",
    "    B -->|是| C[处理]",
    "    B -->|否| D[结束]",
    "    C --> D",
    "```",
  ].join("\n");
  withEditorView((view) => {
    const { state } = view;
    const range = safeSelection(state);
    const lineStart = state.doc.lineAt(range.from).from;
    const before = state.sliceDoc(Math.max(0, lineStart - 2), lineStart);
    const prefix = lineStart === 0 || before.endsWith("\n\n") ? "" : "\n";
    const insert = `${prefix}${template}\n`;
    view.dispatch({
      changes: safeChange(state, range.from, range.to, insert),
      selection: { anchor: range.from + insert.length },
      scrollIntoView: true,
    });
    view.focus();
  });
}

/** 复制选区文本到剪贴板 */
export async function copySelection(): Promise<boolean> {
  const view = getEditorView();
  if (!view) return false;
  const { from, to } = safeSelection(view.state);
  const text = view.state.sliceDoc(from, to);
  if (!text) return false;
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}

/** 剪切选区文本 */
export async function cutSelection(): Promise<boolean> {
  const copied = await copySelection();
  if (!copied) return false;
  withEditorView((view) => {
    const state = view.state;
    const { from, to } = safeSelection(state);
    if (from === to) return;
    view.dispatch({
      changes: safeChange(state, from, to, ""),
      selection: { anchor: from },
      scrollIntoView: true,
    });
    view.focus();
  });
  return true;
}

/** 从剪贴板粘贴纯文本（读取失败时返回 false，由调用方提示改用 Ctrl+V） */
export async function pasteFromClipboard(): Promise<boolean> {
  try {
    const text = await navigator.clipboard.readText();
    if (!text) return false;
    insertText(text);
    return true;
  } catch {
    return false;
  }
}

/** 全选编辑器文本 */
export function selectAllText(): void {
  withEditorView((view) => {
    view.dispatch({
      selection: { anchor: 0, head: view.state.doc.length },
    });
    view.focus();
  });
}

/** 清除选区内的内联标记（粗体 / 斜体 / 删除线 / 行内代码 / 上下标） */
export function clearFormatting(): void {
  withEditorView((view) => {
    const state = view.state;
    const { from, to } = safeSelection(state);
    if (from === to) return;
    const selected = state.sliceDoc(from, to);
    const cleaned = selected
      .replace(/\*\*\*(.+?)\*\*\*/g, "$1")
      .replace(/\*\*(.+?)\*\*/g, "$1")
      .replace(/\*(.+?)\*/g, "$1")
      .replace(/~~(.+?)~~/g, "$1")
      .replace(/`(.+?)`/g, "$1")
      .replace(/\^(.+?)\^/g, "$1")
      .replace(/~(.+?)~/g, "$1");
    if (cleaned === selected) return;
    view.dispatch({
      changes: safeChange(state, from, to, cleaned),
      selection: { anchor: from, head: from + cleaned.length },
      scrollIntoView: true,
    });
    view.focus();
  });
}

/** 插入 Markdown 链接（由链接对话框调用） */export function insertMarkdownLink(text: string, url: string): void {
  const label = text.trim() || url.trim();
  const href = url.trim() || "https://";
  insertInlineSnippet(`[${label}](${href})`);
}

/** 插入 Markdown 图片（由链接对话框调用） */
export function insertMarkdownImage(alt: string, src: string): void {
  const label = alt.trim() || "图片";
  const path = src.trim() || "assets/image.png";
  insertInlineSnippet(`![${label}](${path})`);
}

/** 用片段替换当前选区（无选区则插入到光标处） */
function insertInlineSnippet(snippet: string): void {
  withEditorView((view) => {
    const state = view.state;
    const range = safeSelection(state);
    view.dispatch({
      changes: safeChange(state, range.from, range.to, snippet),
      selection: { anchor: range.from + snippet.length },
      scrollIntoView: true,
    });
    view.focus();
  });
}

/** 切换预览中某一行的任务勾选状态（预览区复选框交互） */
export function toggleTaskOnLine(lineIndex: number, docId?: string): void {
  const state = useAppStore.getState();
  const targetId = docId ?? state.activeId;
  const doc = state.docs.find((d) => d.id === targetId);
  if (!doc) return;
  const lines = doc.content.split("\n");
  if (lineIndex < 0 || lineIndex >= lines.length) return;
  const text = lines[lineIndex];
  const match = /^(\s*[-*+]\s+\[)([ xX])(\])/.exec(text);
  if (!match) return;
  const next = `${match[1]}${match[2] === " " ? "x" : " "}${match[3]}${text.slice(match[0].length)}`;

  // 取「预览所属文档」的编辑器：非激活分栏的预览勾选不能写到激活文档上
  const view = getEditorView(doc.id === state.activeId ? undefined : doc.id);
  if (view && state.viewMode !== "preview") {
    // 源码编辑时走 CodeMirror 事务，保留撤销历史
    const line = view.state.doc.line(lineIndex + 1);
    view.dispatch({
      changes: safeChange(view.state, line.from, line.to, next),
    });
    return;
  }
  lines[lineIndex] = next;
  const content = lines.join("\n");
  if (doc.id === state.activeId) {
    useAppStore.getState().setContent(content);
  } else {
    // setContent 只会写激活文档，非激活分栏必须按文档 ID 写回
    useAppStore.getState().patchDoc(doc.id, { content });
  }
}

/** 选中并滚动到指定偏移区间 */
export function selectRange(from: number, to: number, focus = false): void {
  withEditorView((view) => {
    const max = view.state.doc.length;
    const start = Math.min(from, max);
    const end = Math.min(to, max);
    view.dispatch({
      selection: EditorSelection.range(start, end),
      effects: EditorView.scrollIntoView(start, { y: "center" }),
    });
    if (focus) view.focus();
  });
}
