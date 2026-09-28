import { create } from "zustand";
import { invoke } from "@tauri-apps/api/core";
import {
  EditorSelection,
  StateEffect,
  StateField,
  RangeSetBuilder,
  type Extension,
} from "@codemirror/state";
import {
  Decoration,
  type DecorationSet,
  EditorView,
  ViewPlugin,
  type ViewUpdate,
} from "@codemirror/view";
import { useAppStore, getDocById, createDoc, nextDocId } from "../stores/appStore";
import { useExplorerStore } from "../stores/explorerStore";
import { showMessage } from "../stores/dialogStore";
import { getEditorView } from "./editorBridge";
import { fileName, isMarkdownDoc, isPdfDoc, normalizePath } from "./filePath";
import { openPath } from "./fileActions";
import { isTauri } from "./persist";

/* ------------------------------------------------------------------ */
/* 数据模型                                                            */
/* ------------------------------------------------------------------ */

export interface SnippetLine {
  line: number;
  text: string;
}

export interface CodeSymbolLocation {
  docId?: string;
  path: string | null;
  fileName: string;
  /** 1-based 行号 */
  line: number;
  /** 1-based 列号 */
  col: number;
  lineText: string;
  kind: "function" | "method" | "class" | "type" | "variable" | "reference";
  isDefinition: boolean;
  previewLines: SnippetLine[];
}

export interface NavHistoryEntry {
  docId: string;
  filePath: string | null;
  fileName: string;
  line: number;
  col: number;
  symbol?: string;
}

export interface PeekState {
  visible: boolean;
  docId: string;
  mode: "peek-def" | "references";
  symbol: string;
  locations: CodeSymbolLocation[];
  selectedIndex: number;
  loading: boolean;
}

export interface PaneNavHistory {
  backStack: NavHistoryEntry[];
  forwardStack: NavHistoryEntry[];
}

interface CodeNavStore {
  /** 按左栏(0) / 右栏(1) 独立维护的后退与前进历史栈，杜绝双栏互串 */
  historyByPane: Record<0 | 1, PaneNavHistory>;
  /** 当前编辑器光标下的标识符名称 */
  activeSymbol: string | null;
  /** 速览定义 / 引用列表浮层状态 */
  peekState: PeekState | null;

  setActiveSymbol: (symbol: string | null) => void;
  pushHistory: (pane: 0 | 1, entry: NavHistoryEntry) => void;
  clearPaneHistory: (pane: 0 | 1) => void;
  setPeekState: (state: PeekState | null) => void;
  setPeekSelectedIndex: (index: number) => void;
  closePeek: () => void;
}

export function isSameHistoryLocation(a: NavHistoryEntry, b: NavHistoryEntry): boolean {
  const sameFile =
    a.filePath && b.filePath
      ? normalizePath(a.filePath) === normalizePath(b.filePath)
      : a.docId === b.docId;
  return Boolean(sameFile && Math.abs(a.line - b.line) <= 1);
}

function pushUniqueEntry(stack: NavHistoryEntry[], entry: NavHistoryEntry): NavHistoryEntry[] {
  const last = stack[stack.length - 1];
  if (last && isSameHistoryLocation(last, entry)) {
    // 更新最新列号与 docId
    return [...stack.slice(0, -1), entry];
  }
  return [...stack.slice(-49), entry];
}

export const useCodeNavStore = create<CodeNavStore>((set) => ({
  historyByPane: {
    0: { backStack: [], forwardStack: [] },
    1: { backStack: [], forwardStack: [] },
  },
  activeSymbol: null,
  peekState: null,

  setActiveSymbol: (symbol) => set({ activeSymbol: symbol }),

  pushHistory: (pane, entry) =>
    set((s) => {
      const current = s.historyByPane[pane] ?? { backStack: [], forwardStack: [] };
      const nextBack = pushUniqueEntry(current.backStack, entry);
      return {
        historyByPane: {
          ...s.historyByPane,
          [pane]: {
            backStack: nextBack,
            forwardStack: [],
          },
        },
      };
    }),

  clearPaneHistory: (pane) =>
    set((s) => ({
      historyByPane: {
        ...s.historyByPane,
        [pane]: { backStack: [], forwardStack: [] },
      },
    })),

  setPeekState: (peekState) => set({ peekState }),

  setPeekSelectedIndex: (index) =>
    set((s) =>
      s.peekState
        ? {
            peekState: {
              ...s.peekState,
              selectedIndex: Math.max(
                0,
                Math.min(index, Math.max(0, s.peekState.locations.length - 1)),
              ),
            },
          }
        : s,
    ),

  closePeek: () => set({ peekState: null }),
}));

/* ------------------------------------------------------------------ */
/* 语言关键字过滤与标识符提取                                          */
/* ------------------------------------------------------------------ */

const LANGUAGE_KEYWORDS = new Set([
  "if", "else", "elif", "for", "while", "do", "switch", "case", "default",
  "break", "continue", "return", "try", "catch", "except", "finally", "throw", "raise",
  "function", "fn", "func", "def", "class", "struct", "interface", "enum", "trait",
  "impl", "type", "typedef", "namespace", "module", "mod", "package", "import",
  "export", "from", "as", "use", "pub", "private", "protected", "public", "internal",
  "static", "final", "const", "let", "var", "val", "mut", "ref", "readonly",
  "async", "await", "yield", "go", "defer", "chan", "select", "match", "where",
  "new", "delete", "typeof", "instanceof", "sizeof", "in", "of", "is", "not", "and", "or",
  "true", "false", "null", "undefined", "void", "None", "True", "False", "nil",
  "this", "self", "super", "crate", "pass", "with", "lambda", "virtual", "override",
  "abstract", "extends", "implements", "number", "string", "boolean", "any", "never",
  "int", "float", "double", "char", "bool", "usize", "isize", "u8", "u16", "u32", "u64",
  "i8", "i16", "i32", "i64", "f32", "f64", "str", "String", "Vec", "Option", "Result",
]);

function isIdentChar(ch: string): boolean {
  if (!ch) return false;
  const code = ch.charCodeAt(0);
  return (
    (code >= 65 && code <= 90) || // A-Z
    (code >= 97 && code <= 122) || // a-z
    (code >= 48 && code <= 57) || // 0-9
    code === 95 || // _
    code === 36 // $
  );
}

function isIdentStart(ch: string): boolean {
  if (!ch) return false;
  const code = ch.charCodeAt(0);
  return (
    (code >= 65 && code <= 90) ||
    (code >= 97 && code <= 122) ||
    code === 95 ||
    code === 36
  );
}

/** 提取文档指定偏移处的标识符（排除纯数字与保留关键字） */
export function getSymbolAtOffset(
  text: string,
  offset: number,
): { name: string; from: number; to: number } | null {
  if (!text || offset < 0 || offset > text.length) return null;

  let start = offset;
  let end = offset;

  // 如果恰好处在词尾，向左退一格判断
  if (start > 0 && !isIdentChar(text[start]) && isIdentChar(text[start - 1])) {
    start -= 1;
    end = start;
  }

  if (!isIdentChar(text[start])) return null;

  while (start > 0 && isIdentChar(text[start - 1])) {
    start -= 1;
  }
  while (end < text.length && isIdentChar(text[end])) {
    end += 1;
  }

  const name = text.slice(start, end);
  if (name.length < 2 || !isIdentStart(name[0]) || LANGUAGE_KEYWORDS.has(name)) {
    return null;
  }

  return { name, from: start, to: end };
}

/** 从当前编辑器选区或光标处获取目标标识符 */
export function getSymbolFromEditor(view: EditorView | null): string | null {
  if (!view) return null;
  const sel = view.state.selection.main;
  const docText = view.state.doc.toString();
  if (!sel.empty && sel.to - sel.from <= 64) {
    const selected = docText.slice(sel.from, sel.to).trim();
    if (/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(selected) && !LANGUAGE_KEYWORDS.has(selected)) {
      return selected;
    }
  }
  return getSymbolAtOffset(docText, sel.head)?.name ?? null;
}

/* ------------------------------------------------------------------ */
/* 前端内存级代码定义与引用分类器（与 Rust 侧规则保持一致）            */
/* ------------------------------------------------------------------ */

function findWholeWordIndex(line: string, symbol: string): number {
  if (!symbol || line.length < symbol.length) return -1;
  let start = 0;
  while (start <= line.length - symbol.length) {
    const idx = line.indexOf(symbol, start);
    if (idx === -1) return -1;
    const prevOk = idx === 0 || !isIdentChar(line[idx - 1]);
    const nextOk =
      idx + symbol.length >= line.length || !isIdentChar(line[idx + symbol.length]);
    if (prevOk && nextOk) return idx;
    start = idx + 1;
  }
  return -1;
}

function classifyDefinitionLine(
  trimmed: string,
  symbol: string,
  posInTrimmed: number,
): CodeSymbolLocation["kind"] | null {
  if (
    trimmed.startsWith("//") ||
    (trimmed.startsWith("#") && !trimmed.startsWith("#define")) ||
    trimmed.startsWith("/*") ||
    trimmed.startsWith("*") ||
    trimmed.startsWith("--")
  ) {
    return null;
  }

  const before = trimmed.slice(0, posInTrimmed).trimEnd();
  const after = trimmed.slice(posInTrimmed + symbol.length).trimStart();

  if (before.endsWith(".") || before.endsWith("->")) return null;

  const tokens = before.split(/[^A-Za-z0-9_]+/).filter(Boolean);
  const lastToken = tokens[tokens.length - 1] ?? "";

  if (["fn", "def", "func", "function", "sub"].includes(lastToken)) {
    return "function";
  }
  if (
    ["class", "struct", "interface", "trait", "impl", "enum", "record", "protocol"].includes(
      lastToken,
    )
  ) {
    return "class";
  }
  if (["type", "typedef", "namespace", "module", "mod"].includes(lastToken)) {
    return "type";
  }
  if (lastToken === "define" && before.startsWith("#")) {
    return "function";
  }
  if (["const", "let", "var", "static", "val"].includes(lastToken)) {
    if (after.startsWith("=") || after.startsWith(":")) {
      if (after.includes("=>") || after.includes("function")) {
        return "function";
      }
      return "variable";
    }
  }

  // Go 接收器方法：`func (r *Repo) Symbol(`
  if (trimmed.startsWith("func ") && before.endsWith(")") && after.startsWith("(")) {
    return "method";
  }

  // C/C++/Java/C#/TS/JS 类方法或函数声明
  if (after.startsWith("(") || after.startsWith("<")) {
    if (
      [
        "if",
        "for",
        "while",
        "switch",
        "catch",
        "return",
        "throw",
        "new",
        "await",
        "yield",
        "else",
        "case",
        "sizeof",
        "typeof",
        "delete",
      ].includes(lastToken)
    ) {
      return null;
    }
    if (
      before.includes("=") ||
      before.includes("(") ||
      before.endsWith(",") ||
      before.endsWith("!")
    ) {
      return null;
    }
    const cleanLine = trimmed.split("//")[0].trimEnd();
    if (
      cleanLine.endsWith("{") ||
      cleanLine.endsWith(") {") ||
      cleanLine.endsWith("){")
    ) {
      if (before.length > 0 || after.includes(")")) {
        return before.length === 0 ? "method" : "function";
      }
    }
  }

  return null;
}

function lineIndentWidth(line: string): number {
  let width = 0;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === " ") width += 1;
    else if (ch === "\t") width += 4;
    else break;
  }
  return width;
}

/**
 * 从定义行 `defIdx` 开始，精准提取整个函数/结构体的代码行
 * （只显示当前函数本身，不包含函数上方或下方的其它代码）
 */
export function extractFunctionLines(lines: string[], defIdx: number): SnippetLine[] {
  if (defIdx < 0 || defIdx >= lines.length) return [];
  const maxScanEnd = Math.min(lines.length, defIdx + 500);
  const firstTrimmed = lines[defIdx].trim();

  // 1. Python / 缩进型代码块识别（如 `def foo(...):` / `async def foo(...):` / `class Foo:`）
  const isIndentLang =
    firstTrimmed.startsWith("def ") ||
    firstTrimmed.startsWith("async def ") ||
    (firstTrimmed.startsWith("class ") && !firstTrimmed.includes("{"));

  let endIdx = defIdx;

  if (isIndentLang) {
    const baseIndent = lineIndentWidth(lines[defIdx]);
    let sigDone = firstTrimmed.endsWith(":");
    let lastBodyLine = defIdx;

    for (let curIdx = defIdx + 1; curIdx < maxScanEnd; curIdx++) {
      const line = lines[curIdx];
      const t = line.trim();
      if (!sigDone) {
        lastBodyLine = curIdx;
        if (t.endsWith(":")) {
          sigDone = true;
        }
        continue;
      }
      if (!t || t.startsWith("#")) continue;
      const indent = lineIndentWidth(line);
      if (indent > baseIndent) {
        lastBodyLine = curIdx;
      } else {
        break;
      }
    }
    endIdx = lastBodyLine;
  } else {
    // 2. 花括号 `{ ... }` 或语句型语言（Rust / C / C++ / TS / JS / Go / Java / C# 等）
    let braceDepth = 0;
    let parenDepth = 0;
    let foundOpenBrace = false;
    let inBlockComment = false;

    outer: for (let curIdx = defIdx; curIdx < maxScanEnd; curIdx++) {
      const line = lines[curIdx];
      let inString: string | null = null;

      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        const next = line[i + 1];

        if (inBlockComment) {
          if (ch === "*" && next === "/") {
            inBlockComment = false;
            i += 1;
          }
          continue;
        }

        if (inString) {
          if (ch === "\\") {
            i += 1;
            continue;
          }
          if (ch === inString) {
            inString = null;
          }
          continue;
        }

        if (ch === "/" && next === "/") {
          break;
        }
        if (ch === "/" && next === "*") {
          inBlockComment = true;
          i += 1;
          continue;
        }
        if (ch === '"' || ch === "'" || ch === "`") {
          inString = ch;
          continue;
        }

        if (ch === "(") {
          parenDepth += 1;
        } else if (ch === ")") {
          parenDepth = Math.max(0, parenDepth - 1);
        } else if (ch === "{") {
          braceDepth += 1;
          foundOpenBrace = true;
        } else if (ch === "}") {
          if (foundOpenBrace) {
            braceDepth -= 1;
            if (braceDepth <= 0) {
              endIdx = curIdx;
              break outer;
            }
          }
        } else if (ch === ";" && !foundOpenBrace && parenDepth === 0) {
          endIdx = curIdx;
          break outer;
        }
      }

      endIdx = curIdx;
      if (!foundOpenBrace && parenDepth === 0 && curIdx >= defIdx + 1) {
        const nextTrimmed = (lines[curIdx + 1] ?? "").trim();
        if (!nextTrimmed.startsWith("{") && !nextTrimmed.startsWith("where")) {
          break;
        }
      }
    }
  }

  const result: SnippetLine[] = [];
  for (let p = defIdx; p <= endIdx; p++) {
    result.push({
      line: p + 1,
      text: lines[p].slice(0, 400),
    });
  }
  return result;
}

/** 扫描单个内存文档中的定义或全部引用 */
export function scanDocumentForSymbol(
  docId: string,
  filePath: string | null,
  content: string,
  symbol: string,
  onlyDefs: boolean,
): CodeSymbolLocation[] {
  if (!content || !symbol || !content.includes(symbol)) return [];
  const lines = content.split(/\r?\n/);
  const name = filePath ? fileName(filePath) : "未命名文档";
  const results: CodeSymbolLocation[] = [];

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i];
    const colIdx = findWholeWordIndex(rawLine, symbol);
    if (colIdx === -1) continue;

    const trimmed = rawLine.trim();
    const posInTrimmed = findWholeWordIndex(trimmed, symbol);
    const defKind =
      posInTrimmed !== -1 ? classifyDefinitionLine(trimmed, symbol, posInTrimmed) : null;
    const isDef = defKind !== null;

    if (onlyDefs && !isDef) continue;

    const previewLines: SnippetLine[] = isDef ? extractFunctionLines(lines, i) : [];

    results.push({
      docId,
      path: filePath,
      fileName: name,
      line: i + 1,
      col: colIdx + 1,
      lineText: trimmed.slice(0, 240),
      kind: defKind ?? "reference",
      isDefinition: isDef,
      previewLines,
    });
  }

  return results;
}

/* ------------------------------------------------------------------ */
/* 三级混合检索：当前文档 -> 已打开标签页 -> Rust 工作区跨文件扫描     */
/* ------------------------------------------------------------------ */

export async function findSymbolLocations(
  sourceDocId: string,
  symbol: string,
  mode: "defs" | "refs",
): Promise<CodeSymbolLocation[]> {
  const onlyDefs = mode === "defs";
  const state = useAppStore.getState();
  const currentDoc = state.docs.find((d) => d.id === sourceDocId) ?? null;
  const seen = new Set<string>();
  const results: CodeSymbolLocation[] = [];

  const addUnique = (loc: CodeSymbolLocation) => {
    const key = `${loc.path ? normalizePath(loc.path) : loc.docId ?? ""}:${loc.line}`;
    if (seen.has(key)) return;
    seen.add(key);
    results.push(loc);
  };

  // Tier 1: 当前活动文档实时缓冲区（无需保存即可立即定位）
  if (currentDoc && !isPdfDoc(currentDoc)) {
    const localHits = scanDocumentForSymbol(
      currentDoc.id,
      currentDoc.filePath,
      currentDoc.content,
      symbol,
      onlyDefs,
    );
    localHits.forEach(addUnique);
  }

  // Tier 2: 已打开的其它代码标签页
  for (const doc of state.docs) {
    if (doc.id === sourceDocId || isPdfDoc(doc) || isMarkdownDoc(doc)) continue;
    const tabHits = scanDocumentForSymbol(
      doc.id,
      doc.filePath,
      doc.content,
      symbol,
      onlyDefs,
    );
    tabHits.forEach(addUnique);
  }

  // Tier 3: 若为桌面端且有工作区目录或文件路径，调用 Rust 原生高速跨文件扫描
  const scanDir =
    useExplorerStore.getState().rootPath || currentDoc?.filePath || null;
  if (isTauri && scanDir) {
    try {
      const workspaceHits = await invoke<CodeSymbolLocation[]>(
        "search_workspace_symbols",
        {
          dirPath: scanDir,
          symbol,
          mode,
        },
      );
      for (const hit of workspaceHits) {
        // 如果该文件已在标签页打开，关联其 docId
        const openDoc = state.docs.find(
          (d) => d.filePath && hit.path && normalizePath(d.filePath) === normalizePath(hit.path),
        );
        addUnique({
          ...hit,
          docId: openDoc?.id,
        });
      }
    } catch (err) {
      console.warn("[MasterEdit] 跨文件符号扫描异常:", err);
    }
  }

  return results;
}

/* ------------------------------------------------------------------ */
/* 编辑器目标行脉冲高亮动画（Jump Target Flash）                        */
/* ------------------------------------------------------------------ */

export const flashLineEffect = StateEffect.define<number | null>();

const jumpFlashDecoration = Decoration.line({ class: "cm-jump-flash" });

const jumpFlashField = StateField.define<DecorationSet>({
  create() {
    return Decoration.none;
  },
  update(deco, tr) {
    deco = deco.map(tr.changes);
    for (const e of tr.effects) {
      if (e.is(flashLineEffect)) {
        if (e.value === null) {
          return Decoration.none;
        }
        const lineNum = Math.max(1, Math.min(e.value, tr.state.doc.lines));
        const line = tr.state.doc.line(lineNum);
        const builder = new RangeSetBuilder<Decoration>();
        builder.add(line.from, line.from, jumpFlashDecoration);
        return builder.finish();
      }
    }
    return deco;
  },
  provide: (f) => EditorView.decorations.from(f),
});

/** 滚动指定编辑器到目标行/列，并触发脉冲高亮动画 */
export function scrollEditorToLineAndFlash(
  docId: string,
  line1Based: number,
  col1Based = 1,
): void {
  const safeInitLine = Math.max(1, line1Based);
  const safeInitCol = Math.max(1, col1Based);
  // 立即同步 store 中的光标行/列，避免异步挂载期间读取到旧行号
  useAppStore.getState().setDocCursor(docId, safeInitLine, safeInitCol);

  const attempt = (retriesLeft: number) => {
    const view = docId ? getEditorView(docId) : getEditorView();
    if (!view) {
      if (retriesLeft > 0) {
        window.setTimeout(() => attempt(retriesLeft - 1), 40);
      }
      return;
    }
    const safeLine = Math.max(1, Math.min(line1Based, view.state.doc.lines));
    const lineInfo = view.state.doc.line(safeLine);
    const targetPos = Math.min(
      lineInfo.to,
      lineInfo.from + Math.max(0, col1Based - 1),
    );
    const actualCol = targetPos - lineInfo.from + 1;

    useAppStore.getState().setDocCursor(docId, safeLine, actualCol);

    view.requestMeasure();
    view.dispatch({
      selection: EditorSelection.cursor(targetPos),
      effects: [
        EditorView.scrollIntoView(targetPos, { y: "center" }),
        flashLineEffect.of(safeLine),
      ],
    });
    view.focus();

    requestAnimationFrame(() => {
      try {
        view.requestMeasure();
        view.dispatch({
          effects: EditorView.scrollIntoView(targetPos, { y: "center" }),
        });
      } catch {
        /* ignore destroyed view */
      }
    });

    window.setTimeout(() => {
      try {
        view.dispatch({ effects: flashLineEffect.of(null) });
      } catch {
        /* ignore destroyed view */
      }
    }, 1500);
  };

  attempt(12);
}

/* ------------------------------------------------------------------ */
/* 核心交互动作：转到定义 / 速览定义 / 查找引用 / 历史前进后退          */
/* ------------------------------------------------------------------ */

/** 实时获取指定文档的精确光标位置（优先从挂载的 EditorView 读取实时光标） */
export function getAccurateDocPosition(
  docId: string,
  symbol?: string,
): NavHistoryEntry | null {
  const doc = getDocById(docId);
  if (!doc) return null;

  let line = doc.cursorLine || 1;
  let col = doc.cursorCol || 1;

  const view = getEditorView(docId);
  if (view) {
    try {
      const head = view.state.selection.main.head;
      const lineObj = view.state.doc.lineAt(head);
      line = lineObj.number;
      col = head - lineObj.from + 1;
    } catch {
      /* fallback to store cursor */
    }
  }

  return {
    docId: doc.id,
    filePath: doc.filePath,
    fileName: doc.filePath ? fileName(doc.filePath) : "未命名文档",
    line,
    col,
    symbol,
  };
}

/** 在当前文档内跳转到指定行并记录分栏历史（供面包屑符号下拉跳转等调用） */
export function jumpToLineInDoc(
  docId: string,
  line1Based: number,
  col1Based = 1,
  symbol?: string,
): void {
  const doc = getDocById(docId);
  if (!doc) return;
  const pane = (doc.pane ?? 0) as 0 | 1;
  const currentPos = getAccurateDocPosition(docId, symbol);
  if (currentPos && Math.abs(currentPos.line - line1Based) > 1) {
    useCodeNavStore.getState().pushHistory(pane, currentPos);
  }
  scrollEditorToLineAndFlash(docId, line1Based, col1Based);
}

/** 跳转到指定的符号位置（支持当前栏跳转或在另一侧分栏并排打开查看） */
export async function jumpToCodeLocation(
  loc: CodeSymbolLocation,
  options?: { openInSplit?: boolean; fromDocId?: string; symbol?: string },
): Promise<void> {
  const appState = useAppStore.getState();
  const fromDocId = options?.fromDocId ?? appState.activeId;
  const sourceDoc = fromDocId ? getDocById(fromDocId) : null;
  const sourcePane = (sourceDoc?.pane ?? appState.layout.activePane ?? 0) as 0 | 1;
  const targetPane: 0 | 1 = options?.openInSplit
    ? sourcePane === 0
      ? 1
      : 0
    : sourcePane;

  const sourceEntry = fromDocId
    ? getAccurateDocPosition(fromDocId, options?.symbol)
    : null;

  useCodeNavStore.getState().closePeek();

  // ==================== 分栏并排打开模式 (openInSplit = true) ====================
  // 核心保证：
  // 1. 绝不把当前正在阅读的 sourceDoc 从 sourcePane 移走！
  // 2. 绝不污染 sourcePane 的历史栈（因为 sourcePane 仍停留在原调用处未动）！
  // 3. 将历史记录精确归入真正发生位置变化的 targetPane 栈中，保证在左/右栏点击「返回上一位置 / 前进下一位置」均在本栏内自洽回退与前进。
  if (options?.openInSplit) {
    const wasTargetPaneVisible = appState.layout.split || targetPane === 0;
    if (!wasTargetPaneVisible) {
      useCodeNavStore.getState().clearPaneHistory(targetPane);
    }

    const prevTargetDocId = wasTargetPaneVisible
      ? appState.activeIds[targetPane]
      : null;
    const prevTargetDoc = prevTargetDocId ? getDocById(prevTargetDocId) : null;
    const hasRealPrevTarget = Boolean(
      prevTargetDoc && (prevTargetDoc.filePath || prevTargetDoc.content.trim()),
    );
    const prevTargetEntry =
      hasRealPrevTarget && prevTargetDocId
        ? getAccurateDocPosition(prevTargetDocId, options?.symbol)
        : null;

    // 辅助函数：为目标分栏记录跳转前的历史位置
    const recordTargetPaneHistory = (destDocId: string, destFilePath: string | null) => {
      const targetLocEntry: NavHistoryEntry = {
        docId: destDocId,
        filePath: destFilePath,
        fileName: destFilePath ? fileName(destFilePath) : loc.fileName,
        line: loc.line,
        col: loc.col,
      };
      if (prevTargetEntry) {
        if (!isSameHistoryLocation(prevTargetEntry, targetLocEntry)) {
          useCodeNavStore.getState().pushHistory(targetPane, prevTargetEntry);
        }
      } else if (sourceEntry) {
        // 若目标分栏此前为空/首次开启分栏，且从源文档分栏打开同一文件，将调用处行号绑定到目标栏文档压入目标栏后退栈
        const fallbackEntry: NavHistoryEntry = {
          ...sourceEntry,
          docId:
            !sourceEntry.filePath ||
            (destFilePath &&
              normalizePath(sourceEntry.filePath) === normalizePath(destFilePath))
              ? destDocId
              : sourceEntry.docId,
        };
        if (!isSameHistoryLocation(fallbackEntry, targetLocEntry)) {
          useCodeNavStore.getState().pushHistory(targetPane, fallbackEntry);
        }
      }
    };

    // 1. 先检查目标分栏 (targetPane) 中是否已经打开了该文件
    const existingInTargetPane = appState.docs.find(
      (d) =>
        (d.pane ?? 0) === targetPane &&
        d.id !== fromDocId &&
        ((loc.docId && d.id === loc.docId) ||
          (loc.path && d.filePath && normalizePath(d.filePath) === normalizePath(loc.path)) ||
          (!loc.path && !d.filePath && sourceDoc && d.content === sourceDoc.content)),
    );

    if (existingInTargetPane) {
      recordTargetPaneHistory(existingInTargetPane.id, existingInTargetPane.filePath);
      const nextActiveIds: [string | null, string | null] = [...appState.activeIds];
      if (sourceDoc) nextActiveIds[sourcePane] = sourceDoc.id;
      nextActiveIds[targetPane] = existingInTargetPane.id;
      useAppStore.setState((s) => ({
        activeIds: nextActiveIds,
        activeId: existingInTargetPane.id,
        layout: { ...s.layout, split: true, activePane: targetPane },
      }));
      scrollEditorToLineAndFlash(existingInTargetPane.id, loc.line, loc.col);
      return;
    }

    // 2. 检查当前已打开文档中是否有该文件，若有则在 targetPane 创建镜像分栏标签页（绝不抽走 sourcePane 的标签）
    const existingInSourcePane =
      (loc.docId ? getDocById(loc.docId) : null) ??
      (loc.path
        ? appState.docs.find(
            (d) => d.filePath && normalizePath(d.filePath) === normalizePath(loc.path!),
          ) ?? null
        : null);

    if (existingInSourcePane) {
      const companionDoc = createDoc({
        ...existingInSourcePane,
        id: nextDocId(),
        pane: targetPane,
        cursorLine: loc.line,
        cursorCol: loc.col,
        scrollTop: 0,
      });
      recordTargetPaneHistory(companionDoc.id, companionDoc.filePath);
      const nextActiveIds: [string | null, string | null] = [...appState.activeIds];
      if (sourceDoc) nextActiveIds[sourcePane] = sourceDoc.id;
      nextActiveIds[targetPane] = companionDoc.id;
      useAppStore.setState((s) => ({
        docs: [...s.docs, companionDoc],
        activeIds: nextActiveIds,
        activeId: companionDoc.id,
        layout: { ...s.layout, split: true, activePane: targetPane },
      }));
      window.setTimeout(() => {
        scrollEditorToLineAndFlash(companionDoc.id, loc.line, loc.col);
      }, 40);
      return;
    }

    // 3. 目标是磁盘上的其它未打开工程文件：开启双栏并在 targetPane 打开该文件
    if (loc.path) {
      const nextActiveIds: [string | null, string | null] = [...appState.activeIds];
      if (sourceDoc) nextActiveIds[sourcePane] = sourceDoc.id;
      useAppStore.setState((s) => ({
        activeIds: nextActiveIds,
        layout: { ...s.layout, split: true, activePane: targetPane },
      }));
      await openPath(loc.path, targetPane);
      const newlyOpened = useAppStore
        .getState()
        .docs.find(
          (d) =>
            (d.pane ?? 0) === targetPane &&
            d.filePath &&
            normalizePath(d.filePath) === normalizePath(loc.path!),
        );
      if (newlyOpened) {
        recordTargetPaneHistory(newlyOpened.id, newlyOpened.filePath);
        window.setTimeout(() => {
          scrollEditorToLineAndFlash(newlyOpened.id, loc.line, loc.col);
        }, 60);
      }
    }
    return;
  }

  // ==================== 当前栏普通跳转模式 (openInSplit = false) ====================
  // 记录当前分栏跳转前的位置到 sourcePane 的后退栈
  if (sourceEntry) {
    const targetCompare: NavHistoryEntry = {
      docId: loc.docId ?? sourceEntry.docId,
      filePath: loc.path ?? sourceEntry.filePath,
      fileName: loc.fileName,
      line: loc.line,
      col: loc.col,
    };
    if (!isSameHistoryLocation(sourceEntry, targetCompare)) {
      useCodeNavStore.getState().pushHistory(sourcePane, sourceEntry);
    }
  }

  // 1. 优先匹配当前分栏 (sourcePane) 中已打开的同文件标签
  const existingInSamePane = appState.docs.find(
    (d) =>
      (d.pane ?? 0) === sourcePane &&
      ((loc.docId && d.id === loc.docId) ||
        (loc.path && d.filePath && normalizePath(d.filePath) === normalizePath(loc.path))),
  );

  if (existingInSamePane) {
    appState.activateDoc(existingInSamePane.id, sourcePane);
    scrollEditorToLineAndFlash(existingInSamePane.id, loc.line, loc.col);
    return;
  }

  // 2. 若当前处于双栏模式且该文件仅在另一栏打开，在当前分栏创建独立标签页，避免跨栏抢夺焦点或移动另一栏标签
  const existingInOtherPane =
    (loc.docId ? getDocById(loc.docId) : null) ??
    (loc.path
      ? appState.docs.find(
          (d) => d.filePath && normalizePath(d.filePath) === normalizePath(loc.path!),
        ) ?? null
      : null);

  if (existingInOtherPane) {
    if (appState.layout.split && (existingInOtherPane.pane ?? 0) !== sourcePane) {
      const companionInCurrentPane = createDoc({
        ...existingInOtherPane,
        id: nextDocId(),
        pane: sourcePane,
        cursorLine: loc.line,
        cursorCol: loc.col,
        scrollTop: 0,
      });
      const nextActiveIds: [string | null, string | null] = [...appState.activeIds];
      nextActiveIds[sourcePane] = companionInCurrentPane.id;
      useAppStore.setState((s) => ({
        docs: [...s.docs, companionInCurrentPane],
        activeIds: nextActiveIds,
        activeId: companionInCurrentPane.id,
        layout: { ...s.layout, activePane: sourcePane },
      }));
      window.setTimeout(() => {
        scrollEditorToLineAndFlash(companionInCurrentPane.id, loc.line, loc.col);
      }, 40);
      return;
    }

    appState.activateDoc(existingInOtherPane.id, sourcePane);
    scrollEditorToLineAndFlash(existingInOtherPane.id, loc.line, loc.col);
    return;
  }

  // 3. 目标是磁盘上的其它未打开工程文件：在当前分栏打开文件后再定位行号
  if (loc.path) {
    await openPath(loc.path, sourcePane);
    const newlyOpened = useAppStore
      .getState()
      .docs.find(
        (d) =>
          (d.pane ?? 0) === sourcePane &&
          d.filePath &&
          normalizePath(d.filePath) === normalizePath(loc.path!),
      );
    if (newlyOpened) {
      window.setTimeout(() => {
        scrollEditorToLineAndFlash(newlyOpened.id, loc.line, loc.col);
      }, 60);
    }
  }
}

/** 转到函数/类/符号的原定义（F12 / Ctrl+Click / 右键菜单） */
export async function goToSymbolDefinition(
  docId: string,
  explicitSymbol?: string,
  options?: { openInSplit?: boolean },
): Promise<void> {
  const view = getEditorView(docId) ?? getEditorView();
  const symbol = explicitSymbol || getSymbolFromEditor(view);
  if (!symbol) {
    await showMessage("未选中符号", "请先将光标停留在函数名、类名或变量名上。");
    return;
  }

  const defs = await findSymbolLocations(docId, symbol, "defs");
  const currentPos = getAccurateDocPosition(docId);
  const currentLine = currentPos?.line ?? 1;

  if (defs.length === 0) {
    // 若未识别出显式声明语法，自动回退展示该符号在工程中的所有出现位置
    const refs = await findSymbolLocations(docId, symbol, "refs");
    if (refs.length > 0) {
      const otherRef = refs.find((r) => !(r.docId === docId && r.line === currentLine));
      if (refs.length === 1 && otherRef) {
        await jumpToCodeLocation(otherRef, { ...options, fromDocId: docId, symbol });
        return;
      }
      useCodeNavStore.getState().setPeekState({
        visible: true,
        docId,
        mode: "references",
        symbol,
        locations: refs,
        selectedIndex: 0,
        loading: false,
      });
      return;
    }
    await showMessage(
      "未找到符号定义",
      `未在当前文件或工程目录中找到「${symbol}」的函数或类型定义（可能是内置 API 或第三方库符号）。`,
    );
    return;
  }

  // 如果光标正好已经在唯一的定义行上，直接弹出「查找所有引用」方便查看何处调用了此函数
  if (
    defs.length === 1 &&
    defs[0].docId === docId &&
    defs[0].line === currentLine &&
    !options?.openInSplit
  ) {
    await openSymbolReferences(docId, symbol);
    return;
  }

  // 优先选择非当前行的定义
  const targetDef =
    defs.find((d) => !(d.docId === docId && d.line === currentLine)) ?? defs[0];

  // 如果存在多处不同文件/不同行的同名定义，且未按住分栏快捷键，优先跳转到首个定义
  await jumpToCodeLocation(targetDef, { ...options, fromDocId: docId, symbol });
}

/** 速览原函数定义（Alt+F12 / 右键「速览原函数定义」）：内联浮窗查看原函数代码 */
export async function peekSymbolDefinition(
  docId: string,
  explicitSymbol?: string,
): Promise<void> {
  const view = getEditorView(docId) ?? getEditorView();
  const symbol = explicitSymbol || getSymbolFromEditor(view);
  if (!symbol) return;

  useCodeNavStore.getState().setPeekState({
    visible: true,
    docId,
    mode: "peek-def",
    symbol,
    locations: [],
    selectedIndex: 0,
    loading: true,
  });

  const defs = await findSymbolLocations(docId, symbol, "defs");
  if (defs.length === 0) {
    const refs = await findSymbolLocations(docId, symbol, "refs");
    useCodeNavStore.getState().setPeekState({
      visible: true,
      docId,
      mode: refs.length > 0 ? "references" : "peek-def",
      symbol,
      locations: refs,
      selectedIndex: 0,
      loading: false,
    });
    return;
  }

  useCodeNavStore.getState().setPeekState({
    visible: true,
    docId,
    mode: "peek-def",
    symbol,
    locations: defs,
    selectedIndex: 0,
    loading: false,
  });
}

/** 查找所有引用（Shift+F12 / 右键「查找所有引用」） */
export async function openSymbolReferences(
  docId: string,
  explicitSymbol?: string,
): Promise<void> {
  const view = getEditorView(docId) ?? getEditorView();
  const symbol = explicitSymbol || getSymbolFromEditor(view);
  if (!symbol) return;

  useCodeNavStore.getState().setPeekState({
    visible: true,
    docId,
    mode: "references",
    symbol,
    locations: [],
    selectedIndex: 0,
    loading: true,
  });

  const refs = await findSymbolLocations(docId, symbol, "refs");
  useCodeNavStore.getState().setPeekState({
    visible: true,
    docId,
    mode: "references",
    symbol,
    locations: refs,
    selectedIndex: 0,
    loading: false,
  });
}

/** 导航历史后退（Alt + ← 或鼠标后退侧键，支持按指定分栏独立后退） */
export async function navigateHistoryBack(explicitPane?: 0 | 1): Promise<boolean> {
  const appState = useAppStore.getState();
  const pane: 0 | 1 =
    explicitPane ?? (appState.layout.split ? appState.layout.activePane : 0);
  const nav = useCodeNavStore.getState();
  const paneHist = nav.historyByPane[pane];
  if (!paneHist || paneHist.backStack.length === 0) return false;

  const activeDocIdInPane = appState.layout.split
    ? appState.activeIds[pane] ?? appState.activeId
    : appState.activeId;
  const currentEntry = activeDocIdInPane
    ? getAccurateDocPosition(activeDocIdInPane)
    : null;

  const workingBack = [...paneHist.backStack];
  let target = workingBack.pop()!;
  while (
    workingBack.length > 0 &&
    currentEntry &&
    isSameHistoryLocation(target, currentEntry)
  ) {
    target = workingBack.pop()!;
  }

  if (currentEntry && isSameHistoryLocation(target, currentEntry)) {
    useCodeNavStore.setState((s) => ({
      historyByPane: {
        ...s.historyByPane,
        [pane]: {
          ...s.historyByPane[pane],
          backStack: workingBack,
        },
      },
    }));
    return false;
  }

  const nextForward = currentEntry
    ? pushUniqueEntry(paneHist.forwardStack, currentEntry)
    : paneHist.forwardStack;

  useCodeNavStore.setState((s) => ({
    historyByPane: {
      ...s.historyByPane,
      [pane]: {
        backStack: workingBack,
        forwardStack: nextForward,
      },
    },
  }));

  await restoreHistoryEntry(target, pane);
  return true;
}

/** 导航历史前进（Alt + → 或鼠标前进侧键，支持按指定分栏独立前进） */
export async function navigateHistoryForward(explicitPane?: 0 | 1): Promise<boolean> {
  const appState = useAppStore.getState();
  const pane: 0 | 1 =
    explicitPane ?? (appState.layout.split ? appState.layout.activePane : 0);
  const nav = useCodeNavStore.getState();
  const paneHist = nav.historyByPane[pane];
  if (!paneHist || paneHist.forwardStack.length === 0) return false;

  const activeDocIdInPane = appState.layout.split
    ? appState.activeIds[pane] ?? appState.activeId
    : appState.activeId;
  const currentEntry = activeDocIdInPane
    ? getAccurateDocPosition(activeDocIdInPane)
    : null;

  const workingForward = [...paneHist.forwardStack];
  let target = workingForward.pop()!;
  while (
    workingForward.length > 0 &&
    currentEntry &&
    isSameHistoryLocation(target, currentEntry)
  ) {
    target = workingForward.pop()!;
  }

  if (currentEntry && isSameHistoryLocation(target, currentEntry)) {
    useCodeNavStore.setState((s) => ({
      historyByPane: {
        ...s.historyByPane,
        [pane]: {
          ...s.historyByPane[pane],
          forwardStack: workingForward,
        },
      },
    }));
    return false;
  }

  const nextBack = currentEntry
    ? pushUniqueEntry(paneHist.backStack, currentEntry)
    : paneHist.backStack;

  useCodeNavStore.setState((s) => ({
    historyByPane: {
      ...s.historyByPane,
      [pane]: {
        backStack: nextBack,
        forwardStack: workingForward,
      },
    },
  }));

  await restoreHistoryEntry(target, pane);
  return true;
}

async function restoreHistoryEntry(
  entry: NavHistoryEntry,
  pane: 0 | 1,
): Promise<void> {
  const appState = useAppStore.getState();

  // 1. 优先在当前分栏 (pane) 内匹配原 docId 或同路径文档，杜绝跨分栏乱跳
  const exactInPane = appState.docs.find(
    (d) => d.id === entry.docId && (d.pane ?? 0) === pane,
  );
  const samePathInPane =
    exactInPane ??
    (entry.filePath
      ? appState.docs.find(
          (d) =>
            (d.pane ?? 0) === pane &&
            d.filePath &&
            normalizePath(d.filePath) === normalizePath(entry.filePath!),
        ) ?? null
      : null);

  if (samePathInPane) {
    appState.activateDoc(samePathInPane.id, pane);
    scrollEditorToLineAndFlash(samePathInPane.id, entry.line, entry.col);
    return;
  }

  // 2. 若当前为单栏模式，允许直接匹配全局已打开文档
  if (!appState.layout.split) {
    const anyExisting =
      getDocById(entry.docId) ??
      (entry.filePath
        ? appState.docs.find(
            (d) =>
              d.filePath && normalizePath(d.filePath) === normalizePath(entry.filePath!),
          ) ?? null
        : null);
    if (anyExisting) {
      appState.activateDoc(anyExisting.id, 0);
      scrollEditorToLineAndFlash(anyExisting.id, entry.line, entry.col);
      return;
    }
  }

  // 3. 若双栏模式下该文件仅在另一侧分栏打开，在本分栏创建副本标签以保持本分栏独立回退
  const otherPaneDoc =
    getDocById(entry.docId) ??
    (entry.filePath
      ? appState.docs.find(
          (d) =>
            d.filePath && normalizePath(d.filePath) === normalizePath(entry.filePath!),
        ) ?? null
      : null);

  if (otherPaneDoc && appState.layout.split) {
    const companion = createDoc({
      ...otherPaneDoc,
      id: nextDocId(),
      pane,
      cursorLine: entry.line,
      cursorCol: entry.col,
      scrollTop: 0,
    });
    const nextActiveIds: [string | null, string | null] = [...appState.activeIds];
    nextActiveIds[pane] = companion.id;
    useAppStore.setState((s) => ({
      docs: [...s.docs, companion],
      activeIds: nextActiveIds,
      activeId: companion.id,
      layout: { ...s.layout, activePane: pane },
    }));
    window.setTimeout(() => {
      scrollEditorToLineAndFlash(companion.id, entry.line, entry.col);
    }, 40);
    return;
  }

  // 4. 从磁盘重新打开到目标分栏
  if (entry.filePath) {
    await openPath(entry.filePath, pane);
    const opened = useAppStore
      .getState()
      .docs.find(
        (d) =>
          (d.pane ?? 0) === pane &&
          d.filePath &&
          normalizePath(d.filePath) === normalizePath(entry.filePath!),
      );
    if (opened) {
      window.setTimeout(() => {
        scrollEditorToLineAndFlash(opened.id, entry.line, entry.col);
      }, 60);
    }
  }
}

/* ------------------------------------------------------------------ */
/* CodeMirror 6 增强扩展：同名符号高亮 + Ctrl+悬停下划线 + Ctrl+单击跳转 */
/* ------------------------------------------------------------------ */

const occurrenceMark = Decoration.mark({ class: "cm-word-occurrence" });
const ctrlLinkMark = Decoration.mark({ class: "cm-ctrl-link" });

const setCtrlHoverRange = StateEffect.define<{ from: number; to: number } | null>();

const ctrlHoverField = StateField.define<DecorationSet>({
  create() {
    return Decoration.none;
  },
  update(deco, tr) {
    deco = deco.map(tr.changes);
    for (const e of tr.effects) {
      if (e.is(setCtrlHoverRange)) {
        if (!e.value || e.value.from >= e.value.to) {
          return Decoration.none;
        }
        const builder = new RangeSetBuilder<Decoration>();
        builder.add(e.value.from, e.value.to, ctrlLinkMark);
        return builder.finish();
      }
    }
    return deco;
  },
  provide: (f) => EditorView.decorations.from(f),
});

/** 光标所在标识符的视口内全词同名高亮插件 */
const occurrenceHighlighterPlugin = ViewPlugin.fromClass(
  class {
    decorations: DecorationSet = Decoration.none;

    constructor(view: EditorView) {
      this.decorations = this.compute(view);
    }

    update(update: ViewUpdate) {
      if (update.docChanged || update.selectionSet || update.viewportChanged) {
        this.decorations = this.compute(update.view);
      }
    }

    compute(view: EditorView): DecorationSet {
      const docText = view.state.doc.toString();
      const head = view.state.selection.main.head;
      const sym = getSymbolAtOffset(docText, head);
      useCodeNavStore.getState().setActiveSymbol(sym?.name ?? null);

      if (!sym || sym.name.length < 2) {
        return Decoration.none;
      }

      const builder = new RangeSetBuilder<Decoration>();
      const word = sym.name;
      for (const { from, to } of view.visibleRanges) {
        const slice = docText.slice(from, to);
        let searchPos = 0;
        while (searchPos <= slice.length - word.length) {
          const idx = slice.indexOf(word, searchPos);
          if (idx === -1) break;
          const absFrom = from + idx;
          const absTo = absFrom + word.length;
          const prevOk = absFrom === 0 || !isIdentChar(docText[absFrom - 1]);
          const nextOk = absTo >= docText.length || !isIdentChar(docText[absTo]);
          if (prevOk && nextOk) {
            builder.add(absFrom, absTo, occurrenceMark);
          }
          searchPos = idx + word.length;
        }
      }
      return builder.finish();
    }
  },
  {
    decorations: (v) => v.decorations,
  },
);

/** 为代码编辑器创建全套符号导航扩展（Ctrl+Click、同名高亮、跳转闪烁等） */
export function createCodeNavigationExtensions(docId: string): Extension[] {
  return [
    jumpFlashField,
    ctrlHoverField,
    occurrenceHighlighterPlugin,
    EditorView.domEventHandlers({
      mousemove: (event, view) => {
        const doc = getDocById(docId);
        if (!doc || isMarkdownDoc(doc)) return false;

        if (!(event.ctrlKey || event.metaKey)) {
          if (view.state.field(ctrlHoverField, false) !== Decoration.none) {
            view.dispatch({ effects: setCtrlHoverRange.of(null) });
          }
          return false;
        }

        const pos = view.posAtCoords({ x: event.clientX, y: event.clientY });
        if (pos === null) {
          view.dispatch({ effects: setCtrlHoverRange.of(null) });
          return false;
        }
        const sym = getSymbolAtOffset(view.state.doc.toString(), pos);
        view.dispatch({
          effects: setCtrlHoverRange.of(sym ? { from: sym.from, to: sym.to } : null),
        });
        return false;
      },
      mouseleave: (_event, view) => {
        if (view.state.field(ctrlHoverField, false) !== Decoration.none) {
          view.dispatch({ effects: setCtrlHoverRange.of(null) });
        }
        return false;
      },
      mousedown: (event, view) => {
        const doc = getDocById(docId);
        if (!doc || isMarkdownDoc(doc)) return false;
        const pane = (doc.pane ?? 0) as 0 | 1;

        // 鼠标侧键支持：后退 (button 3) / 前进 (button 4)，严格作用于当前编辑器所在分栏
        if (event.button === 3) {
          event.preventDefault();
          void navigateHistoryBack(pane);
          return true;
        }
        if (event.button === 4) {
          event.preventDefault();
          void navigateHistoryForward(pane);
          return true;
        }

        // Ctrl + 左键单击：直接转到定义（Ctrl + Alt + 左键单击：在另一侧分栏打开定义）
        if (event.button === 0 && (event.ctrlKey || event.metaKey)) {
          const pos = view.posAtCoords({ x: event.clientX, y: event.clientY });
          if (pos !== null) {
            const sym = getSymbolAtOffset(view.state.doc.toString(), pos);
            if (sym) {
              event.preventDefault();
              view.dispatch({
                selection: EditorSelection.cursor(pos),
                effects: setCtrlHoverRange.of(null),
              });
              void goToSymbolDefinition(docId, sym.name, {
                openInSplit: event.altKey,
              });
              return true;
            }
          }
        }
        return false;
      },
    }),
  ];
}
