export const APP_NAME = "MasterEdit";
/** 作者信息 */
export const AUTHOR_NAME = "Master Wang";
export const AUTHOR_NAME_CN = "王大师";
export const PROJECT_URL = "https://github.com/bigmasterpang/MasterEdit";

/** 可打开为 Markdown 的扩展名 */
export const MARKDOWN_EXTENSIONS = ["md", "markdown", "mdown", "mkd", "mkdn"];

/** 纯文本扩展名（以源码模式打开，无语法高亮） */
export const TEXT_EXTENSIONS = ["txt", "text", "log", "nfo", "me"];

/** 代码 / 配置类扩展名（按扩展名匹配语法高亮） */
export const CODE_EXTENSIONS = [
  // 配置与数据
  "json", "jsonc", "json5", "yml", "yaml", "toml", "ini", "cfg", "conf", "properties", "env",
  "csv", "tsv", "xml", "plist", "reg",
  // Web
  "html", "htm", "xhtml", "css", "scss", "less", "sass", "js", "mjs", "cjs", "jsx", "ts", "tsx",
  "vue", "svelte", "astro",
  // 脚本
  "py", "rb", "php", "pl", "pm", "lua", "sh", "bash", "zsh", "bat", "cmd", "ps1", "psm1",
  "groovy", "r", "jl", "tcl", "ahk", "vbs",
  // 编译型语言
  "c", "h", "cc", "cpp", "cxx", "hpp", "hh", "cs", "java", "kt", "kts", "scala", "go", "rs",
  "swift", "m", "mm", "dart", "pas", "f90", "f95", "asm", "s", "vb", "fs", "hs", "erl", "ex",
  "clj", "lisp", "el", "nim", "cr", "zig",
  // 数据与其它
  "sql", "proto", "graphql", "gql", "diff", "patch", "tex", "bib", "cmake", "gradle", "mk",
  "dockerfile", "makefile", "gitignore", "editorconfig", "nginx", "service", "srt", "vtt",
];

/** PDF 扩展名 */
export const PDF_EXTENSIONS = ["pdf"];

/**
 * 电子表格扩展名：由 Rust 侧 calamine 解析，**只读**查看
 * （xlsx / xls / xlsb / ods，不注册为系统默认打开程序，避免劫持 Office 关联）
 */
export const SPREADSHEET_EXTENSIONS = ["xlsx", "xls", "xlsb", "ods"];

/** 分隔符文本表格扩展名：走现有文本管线（编码检测、保存、撤销都复用），额外提供表格视图 */
export const DELIMITED_EXTENSIONS = ["csv", "tsv"];

export const OPENABLE_EXTENSIONS = [
  ...MARKDOWN_EXTENSIONS,
  ...PDF_EXTENSIONS,
  ...SPREADSHEET_EXTENSIONS,
  ...TEXT_EXTENSIONS,
  ...CODE_EXTENSIONS,
];

/** 打开对话框过滤器（精简分类显示，避免单个选项因扩展名过多导致下拉列表超长） */
export const OPEN_DIALOG_FILTERS = [
  { name: "常用支持文件", extensions: ["md", "markdown", "pdf", "xlsx", "csv", "txt", "json", "js", "ts", "html", "py", "rs"] },
  { name: "Markdown 文档 (*.md)", extensions: ["md", "markdown", "mdown"] },
  { name: "PDF 文档 (*.pdf)", extensions: ["pdf"] },
  { name: "Excel 表格 (*.xlsx, *.xls, *.ods)", extensions: ["xlsx", "xls", "xlsb", "ods"] },
  { name: "CSV 表格 (*.csv, *.tsv)", extensions: ["csv", "tsv"] },
  { name: "纯文本文档 (*.txt)", extensions: ["txt", "log", "ini", "conf"] },
  {
    name: "常见代码与数据",
    extensions: ["json", "js", "ts", "py", "rs", "go", "java", "c", "cpp", "yaml", "yml", "toml", "sql", "sh", "bat", "ps1"],
  },
  { name: "所有文件 (*.*)", extensions: ["*"] },
];

/** 支持的文件编码（与 Rust 端保持一致） */
export const ENCODINGS = [
  { id: "utf-8", label: "UTF-8" },
  { id: "utf-8-bom", label: "UTF-8 BOM" },
  { id: "gb18030", label: "GB18030 / GBK" },
  { id: "big5", label: "Big5 繁体" },
  { id: "shift_jis", label: "Shift-JIS 日文" },
  { id: "euc-kr", label: "EUC-KR 韩文" },
  { id: "utf-16le", label: "UTF-16 LE" },
  { id: "utf-16be", label: "UTF-16 BE" },
  { id: "windows-1252", label: "Latin-1 (1252)" },
] as const;

export type EncodingId = (typeof ENCODINGS)[number]["id"];

export const EOL_OPTIONS = [
  { id: "lf", label: "LF" },
  { id: "crlf", label: "CRLF" },
  { id: "cr", label: "CR" },
] as const;

export type EolId = (typeof EOL_OPTIONS)[number]["id"];

/** 超过该大小提示只读打开 */
/** 超过 10MB 的文件以只读方式打开（表格类见下） */
export const LARGE_FILE_BYTES = 10 * 1024 * 1024;

/**
 * CSV / TSV 的自动只读阈值：表格视图走惰性行索引（只扫行首偏移、按需解析行），
 * 十几 MB / 十几万行也能流畅编辑，因此放宽到 64MB；再大才默认只读。
 */
export const DELIMITED_EDIT_LIMIT = 64 * 1024 * 1024;

/** 超过该大小关闭实时预览（分屏改为手动刷新） */
export const REALTIME_PREVIEW_LIMIT = 1024 * 1024;

/** 超过该行数降低滚动同步频率 */
export const HEAVY_DOC_LINES = 5000;

/** 预览渲染防抖（毫秒） */
export const PREVIEW_DEBOUNCE = 150;

/** 大纲解析防抖（毫秒） */
export const OUTLINE_DEBOUNCE = 300;

/** 自动保存最短间隔（秒） */
export const MIN_AUTOSAVE_INTERVAL = 5;

export const MIN_FONT_SIZE = 11;
export const MAX_FONT_SIZE = 24;


/** 空文档默认内容 */
export const EMPTY_DOC_PLACEHOLDER = `# 未命名文档

开始输入 Markdown 内容……

- 使用 \`Ctrl+S\` 保存
- 使用 \`Ctrl+E\` 切换视图模式
`;
