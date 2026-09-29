import {
  CODE_EXTENSIONS,
  DELIMITED_EDIT_LIMIT,
  DELIMITED_EXTENSIONS,
  LARGE_FILE_BYTES,
  MARKDOWN_EXTENSIONS,
  OPENABLE_EXTENSIONS,
  SPREADSHEET_EXTENSIONS,
  TEXT_EXTENSIONS,
} from "./constants";

/** 文档类别：Markdown（预览/大纲/导出）/ 代码（语法高亮）/ 纯文本 / PDF / 表格（CSV、Excel） */
export type DocKind = "markdown" | "text" | "code" | "pdf" | "delimited" | "spreadsheet";

/** 统一使用反斜杠，便于 Windows 路径比较 */
export function normalizeSlashes(p: string): string {
  return p.replace(/[\\/]+/g, "\\");
}

/** 规范化路径（统一斜杠并转为小写），用于跨文件比对 */
export function normalizePath(p: string): string {
  return normalizeSlashes(p).toLowerCase();
}

export function extName(p: string): string {
  const base = p.replace(/[\\/]+$/, "").split(/[\\/]/).pop() ?? "";
  const idx = base.lastIndexOf(".");
  return idx > 0 ? base.slice(idx + 1).toLowerCase() : "";
}

export function fileName(p: string | null | undefined): string {
  if (!p) return "未命名";
  const base = p.split(/[\\/]/).pop();
  return base && base.length > 0 ? base : p;
}

export function dirName(p: string): string {
  const idx = Math.max(p.lastIndexOf("\\"), p.lastIndexOf("/"));
  return idx > 0 ? p.slice(0, idx) : "";
}

export function isMarkdownPath(p: string): boolean {
  return MARKDOWN_EXTENSIONS.includes(extName(p));
}

export function isPdfPath(p: string | null | undefined): boolean {
  if (!p) return false;
  return extName(p) === "pdf";
}

/** 判断文档是否为 PDF 文档 */
export function isPdfDoc(
  doc: { filePath: string | null; docType?: "markdown" | "blank" | "pdf" | "spreadsheet" } | null | undefined,
): boolean {
  if (!doc) return false;
  if (doc.docType === "pdf") return true;
  if (doc.filePath && isPdfPath(doc.filePath)) return true;
  return false;
}

/** 是否为电子表格文件（xlsx / xls / xlsb / ods，只读查看） */
export function isSpreadsheetPath(p: string | null | undefined): boolean {
  if (!p) return false;
  return SPREADSHEET_EXTENSIONS.includes(extName(p));
}

/** 是否为分隔符文本表格（csv / tsv，走文本管线 + 表格视图） */
export function isDelimitedPath(p: string | null | undefined): boolean {
  if (!p) return false;
  return DELIMITED_EXTENSIONS.includes(extName(p));
}

/** 判断文档是否为电子表格（只读，由 Rust 侧解析） */
export function isSpreadsheetDoc(
  doc: { filePath: string | null; docType?: "markdown" | "blank" | "pdf" | "spreadsheet" } | null | undefined,
): boolean {
  if (!doc) return false;
  if (doc.docType === "spreadsheet") return true;
  return isSpreadsheetPath(doc.filePath);
}

/** 判断文档是否为 CSV/TSV（表格视图 + 源码编辑） */
export function isDelimitedDoc(
  doc: { filePath: string | null; docType?: "markdown" | "blank" | "pdf" | "spreadsheet" } | null | undefined,
): boolean {
  if (!doc) return false;
  if (doc.docType === "pdf" || doc.docType === "spreadsheet") return false;
  return isDelimitedPath(doc.filePath);
}

/**
 * 判断文档是否应作为 Markdown 处理。
 * - 未保存的空白文档（docType === 'blank' 且无 filePath）明确不是 Markdown。
 * - PDF 文档明确不是 Markdown。
 * - 有路径的文件根据扩展名判断。
 * - 其余未保存文档默认为 Markdown。
 */
export function isMarkdownDoc(
  doc: { filePath: string | null; docType?: "markdown" | "blank" | "pdf" | "spreadsheet" } | null | undefined,
): boolean {
  if (!doc) return false;
  if (doc.docType === "blank" && !doc.filePath) return false;
  if (doc.docType === "pdf" || doc.docType === "spreadsheet") return false;
  if (doc.filePath) {
    if (isPdfPath(doc.filePath)) return false;
    if (isSpreadsheetPath(doc.filePath)) return false;
    return isMarkdownPath(doc.filePath);
  }
  return doc.docType !== "blank";
}

/** 获取文档基础名称（不含同名编号） */
export function getDocBaseName(
  doc: { filePath: string | null; docType?: "markdown" | "blank" | "pdf" | "spreadsheet" },
): string {
  if (doc.filePath) return fileName(doc.filePath);
  if (doc.docType === "blank") return "未命名";
  if (doc.docType === "pdf") return "未命名.pdf";
  return "未命名.md";
}

/** 获取文档展示标题。如果有多个同名文档，自动加上编号区分，如 README.md (1), README.md (2) */
export function getDocTitle(
  doc: { id: string; filePath: string | null; docType?: "markdown" | "blank" | "pdf" | "spreadsheet" },
  allDocs: Array<{ id: string; filePath: string | null; docType?: "markdown" | "blank" | "pdf" | "spreadsheet" }>,
): string {
  const base = getDocBaseName(doc);
  const duplicates = allDocs.filter((d) => getDocBaseName(d) === base);
  if (duplicates.length <= 1) return base;
  const idx = duplicates.findIndex((d) => d.id === doc.id);
  const num = idx >= 0 ? idx + 1 : 1;
  return `${base} (${num})`;
}

/** 文档类别（未保存的新文档按 Markdown 处理） */
export function docKindOf(p: string | null | undefined): DocKind {
  if (!p) return "markdown";
  const ext = extName(p);
  if (ext === "pdf") return "pdf";
  if (SPREADSHEET_EXTENSIONS.includes(ext)) return "spreadsheet";
  if (DELIMITED_EXTENSIONS.includes(ext)) return "delimited";
  if (MARKDOWN_EXTENSIONS.includes(ext)) return "markdown";
  if (CODE_EXTENSIONS.includes(ext)) return "code";
  if (TEXT_EXTENSIONS.includes(ext)) return "text";
  // 未登记扩展名：有扩展名按代码处理（尝试高亮），无扩展名按纯文本
  return ext === "" ? "text" : "code";
}


/**
 * 文本类文档「自动只读」的体积阈值。
 * CSV / TSV 的表格视图是惰性解析，编辑成本与体积基本无关，因此阈值放宽（见 DELIMITED_EDIT_LIMIT）；
 * 其它文本走 CodeMirror，仍按 10MB 处理。
 */
export function textReadOnlyLimit(p: string | null | undefined): number {
  return isDelimitedPath(p) ? DELIMITED_EDIT_LIMIT : LARGE_FILE_BYTES;
}

export function isOpenablePath(p: string): boolean {
  const ext = extName(p);
  return ext === "" ? false : OPENABLE_EXTENSIONS.includes(ext);
}

/** Windows 路径比较：忽略大小写与分隔符差异 */
export function samePath(a: string | null, b: string | null): boolean {
  if (!a || !b) return false;
  return normalizeSlashes(a).toLowerCase() === normalizeSlashes(b).toLowerCase();
}

/** 是否是外部 URL（http/https/mailto/data/asset 等） */
export function isExternalUrl(src: string): boolean {
  return /^(https?:|mailto:|tel:|data:|asset:|blob:|#)/i.test(src.trim());
}


/** 拼接目录与相对路径 */
export function joinPath(dir: string, rel: string): string {
  if (!dir) return rel;
  const d = dir.replace(/[\\/]+$/, "");
  const r = rel.replace(/^[\\/]+/, "");
  return `${d}\\${r.replace(/\//g, "\\")}`;
}

/**
 * 把 Markdown 中的图片地址解析为本地绝对路径。
 * 已保存文档的相对路径基于文档目录解析；未保存文档返回 null（无法解析）。
 */
export function resolveLocalImagePath(
  docPath: string | null,
  src: string,
): string | null {
  const raw = src.trim();
  if (!raw || isExternalUrl(raw)) return null;
  let decoded = raw;
  try {
    decoded = decodeURIComponent(raw);
  } catch {
    decoded = raw;
  }
  // 去掉查询串与锚点
  decoded = decoded.split("#")[0].split("?")[0];
  if (!decoded) return null;
  if (/^[a-zA-Z]:[\\/]/.test(decoded) || decoded.startsWith("\\\\")) {
    return decoded.replace(/\//g, "\\");
  }
  if (!docPath) return null;
  const dir = dirName(docPath);
  if (!dir) return null;
  return joinPath(dir, decoded);
}


