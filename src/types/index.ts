/** 视图模式：预览 / 源码 / 分屏 */
export type ViewMode = "preview" | "source" | "split";

/** 主题模式 */
export type ThemeMode = "light" | "dark" | "system";

/** 配色方案 */
export type AccentName = "blue" | "violet" | "emerald" | "amber" | "rose";

/** 后端 read_markdown_file 返回 */
export interface FilePayload {
  path: string;
  content: string;
  modifiedAt: number;
  size: number;
  /** 是否为企业透明加密文档（已自动解密） */
  encrypted?: boolean;
  /** 加密文档的 4096 字节文件头（base64），保存时用于按原格式加密写回 */
  encryptedHeader?: string | null;
  /** 检测（或指定）的文件编码 */
  encoding?: string;
  /** 检测到的换行符：lf / crlf / cr */
  eol?: string;
}

/** 后端 file-changed 事件载荷 */
export interface FileChangedPayload {
  path: string;
  exists: boolean;
  modifiedAt: number;
  size: number;
}

/** 大纲条目 */
export interface HeadingItem {
  level: number;
  text: string;
  id: string;
  /** 在源码中的行号（0 起） */
  line: number;
}

/** 后端 read_binary_file 返回 */
export interface BinaryPayload {
  path: string;
  dataBase64: string;
  modifiedAt: number;
  size: number;
  /** 是否为企业透明加密文档（已自动解密） */
  encrypted?: boolean;
  /** 加密文档的 4096 字节文件头（base64），保存时用于按原格式加密写回 */
  encryptedHeader?: string | null;
}

/* ------------------------------------------------------------------ */
/* 办公文档（电子表格）：xlsx / xls / ods 由 Rust 侧 calamine 解析只读   */
/* ------------------------------------------------------------------ */

/** 单元格类型：用于对齐与样式（空 / 文本 / 数字 / 布尔 / 日期 / 错误） */
export type SheetCellType = "empty" | "text" | "number" | "bool" | "date" | "error";

/** 单个单元格（`v` 为可直接显示的文本；`f` 为公式，仅含公式的单元格才有） */
export interface SheetCell {
  v: string;
  t: SheetCellType;
  f?: string;
}

/** 工作表元信息（行列为 0 表示尚未加载：打开该表时才知道真实尺寸） */
export interface SheetMeta {
  name: string;
}

/** 后端 spreadsheet_info 返回：文件信息 + 工作表列表 */
export interface SpreadsheetInfo {
  path: string;
  sheets: SheetMeta[];
  modifiedAt: number;
  size: number;
  /** 是否为企业透明加密文档（已在内存中解密，不落明文副本） */
  encrypted: boolean;
  /** 是否支持写回编辑：只有 .xlsx 可以（xlsm 含宏、xls/xlsb/ods 结构不同，一律只读） */
  editable: boolean;
  /** 是否有未落盘的结构改动（内存影子工作簿与磁盘不一致） */
  pending?: boolean;
}

/** 表格编辑的写回方式 */
export type SheetEditKind = "text" | "number" | "bool" | "formula" | "date" | "empty";

/** 一个待提交的单元格编辑（坐标为 0 起的绝对行列，与网格一致） */
export interface SheetEdit {
  sheet: string;
  row: number;
  col: number;
  kind: SheetEditKind;
  value: string;
  /**
   * 公式的计算结果（仅用于界面显示，不写进文件）。
   * 由后端 spreadsheet_eval 算出；为 undefined 表示还没算或不是公式。
   */
  computed?: string | null;
  /** 公式计算失败的原因（同样只用于显示） */
  error?: string | null;
}

/** 后端 spreadsheet_save 返回 */
export interface SheetSaveResult {
  path: string;
  size: number;
  modifiedAt: number;
  /** 实际写入的单元格数 */
  savedCells: number;
  /** 备份文件路径（原文件字节的副本，未备份时为 null） */
  backupPath: string | null;
}

/** 后端 spreadsheet_stats 返回：矩形区域的数值统计（选区超出一屏时用） */
export interface RangeStats {
  cells: number;
  nonEmpty: number;
  numeric: number;
  sum: number;
  average: number | null;
  min: number | null;
  max: number | null;
}

/**
 * 后端 spreadsheet_structure 返回。
 * 结构操作（插入/删除行列、工作表增删改复制）只改**内存影子工作簿**，
 * 不落盘；只有 Ctrl+S（spreadsheet_save）才写回文件。
 */
export interface StructureResult {
  path: string;
  /** 操作后的完整工作表列表（据此刷新底部标签） */
  sheets: SheetMeta[];
  /** 固定 false：结构操作不再直接落盘 */
  saved: boolean;
  /** 是否有未落盘的结构改动 */
  pending: boolean;
}

/** 后端 spreadsheet_state 返回：未落盘状态 + 工作表列表 + 撤销/重做可用性 */
export interface SpreadsheetState {
  path: string;
  sheets: SheetMeta[];
  pending: boolean;
  canUndo: boolean;
  canRedo: boolean;
}

/** 后端 spreadsheet_eval 的请求项（0 起坐标） */
export interface EvalRequest {
  row: number;
  col: number;
  formula: string;
}

/**
 * 后端 spreadsheet_eval 的结果项。
 * 只是**界面显示用**的计算结果，不会写进文件（文件里仍只有公式，Excel 打开时自己算）。
 */
export interface EvalResult {
  row: number;
  col: number;
  /** 计算成功的显示文本 */
  value: string | null;
  /** 计算失败的原因（「暂不支持函数 X」「循环引用」或 #DIV/0! 这类 Excel 错误码） */
  error: string | null;
}
export interface ShadowEditResult {
  path: string;
  sheets: SheetMeta[];
  /** 影子当前是否与磁盘不同 */
  pending: boolean;
  canUndo: boolean;
  canRedo: boolean;
}

/** 后端 spreadsheet_rows 返回：按窗口读取的行数据（配合虚拟滚动） */
export interface SpreadsheetWindow {
  /** 工作表名 */
  sheet: string;
  /** 该表总行数（绝对行号 + 1，与 Excel 行号一致） */
  rows: number;
  /** 该表展示列数（已按上限截断） */
  cols: number;
  /** 本窗口起始行（0 起，绝对行号） */
  start: number;
  /** 本窗口行数据，长度 ≤ 请求的 count；单元格按绝对行列坐标排列 */
  cells: SheetCell[][];
  /** 是否存在因上限而截断的内容 */
  truncated: boolean;
}

/** 单个文档（标签页）状态 */
export interface DocState {
  id: string;
  filePath: string | null;
  content: string;
  savedContent: string;
  isDirty: boolean;
  /** 超过 10MB 的文件以只读方式打开 */
  readOnly: boolean;
  /** 企业加密文档（打开时已解密，保存时按原格式加密写回） */
  encrypted: boolean;
  /** 加密文档的文件头（base64），保存时用于重新加密 */
  encryptedHeader: string | null;
  /** 文件编码（保存时按此编码写回） */
  encoding: string;
  /** 换行符（保存时统一转换为该换行符） */
  eol: string;
  /** 磁盘上的最后修改时间基线（毫秒） */
  modifiedAt: number;
  size: number;
  cursorLine: number;
  cursorCol: number;
  selectionLength: number;
  headings: HeadingItem[];
  frontMatter: Record<string, unknown> | null;
  frontMatterRaw: string | null;
  /** 编辑器滚动位置（切换标签时恢复） */
  scrollTop: number;
  /** 所属分栏：0（左栏/默认），1（右栏） */
  pane: 0 | 1;
  /** 当前文档独立缩放字号（未设置时跟随全局默认字号，互不影响双栏） */
  fontSize?: number;
  /** 新建文档类型：markdown / blank / pdf / spreadsheet（xlsx 可轻量编辑，xls/xlsb/ods 只读查看） */
  docType?: "markdown" | "blank" | "pdf" | "spreadsheet";
  /** 表格待提交的单元格编辑（未保存前只存在于内存，保存时一次性写回） */
  sheetEdits?: SheetEdit[];
  /** 表格首次保存前是否已经确认过「重写工作簿」的提示（每个文档一次） */
  sheetSaveConfirmed?: boolean;
  /** 表格首次结构操作前是否已经确认过「重写工作簿 / 公式不重算」的说明（每个文档一次） */
  sheetStructureConfirmed?: boolean;
  /** 表格是否有未落盘的结构改动（插入/删除行列、工作表增删改复制只在内存中，Ctrl+S 才写回） */
  sheetStructurePending?: boolean;
  /**
   * 手动调整过的列宽 / 行高（列/行索引 → 像素）。
   * 由父组件持有，因此能进同一条撤销栈（Ctrl+Z 可撤回调整）；目前只在会话内有效，不写回文件。
   */
  sheetColumnWidths?: Record<number, number>;
  sheetRowHeights?: Record<number, number>;
  /** 冻结的行数 / 列数（0 = 不冻结；1 = 冻结首行/首列），滚动时始终可见 */
  sheetFreezeRows?: number;
  sheetFreezeCols?: number;
  /** PDF 文件的二进制数据（Base64 编码，编辑如删页/旋转后会更新并置 isDirty） */
  pdfBase64?: string;
  /** PDF 原始/已保存的二进制数据（Base64），用于判断脏状态或恢复 */
  savedPdfBase64?: string;
  /** PDF 当前页码（1-based） */
  pdfCurrentPage?: number;
  /** PDF 总页数 */
  pdfTotalPages?: number;
  /** PDF 缩放比例（例如 1.0, 1.25, 1.5, 或 "width", "page"） */
  pdfScale?: number | "width" | "page";
  /** PDF 密码（如果是密码加密文件） */
  pdfPassword?: string;
  /** PDF 基础纯净数据（Base64，未绘制可撤销高亮） */
  cleanPdfBase64?: string;
  /** PDF 动态高亮标注列表（可新增、删除、清除） */
  pdfHighlights?: PdfHighlight[];
  /** PDF 阅读底色主题：white, warm, green, parchment, dark */
  pdfPaperTheme?: string;
  /** PDF 便签附注标注列表 */
  pdfNotes?: PdfNote[];
}

/** PDF 便签附注图钉数据 */
export interface PdfNote {
  id: string;
  page: number;
  xPercent: number;
  yPercent: number;
  content: string;
  color?: "yellow" | "blue" | "green" | "purple";
  createdAt: number;
}

/** PDF 动态高亮矩形标注 */
export interface PdfHighlight {
  id: string;
  page: number;
  rects: Array<{
    xPercent: number;
    yPercent: number;
    wPercent: number;
    hPercent: number;
  }>;
  color: "yellow" | "green" | "pink";
  text?: string;
  comment?: string;
  createdAt: number;
}

/** 双栏文档布局状态 */
export interface LayoutState {
  split: boolean;
  activePane: 0 | 1;
  ratio: number;
}

/** 可自定义的快捷键 */
export type ShortcutId =
  | "open"
  | "save"
  | "saveAs"
  | "newDoc"
  | "viewMode"
  | "search"
  | "replace"
  | "closeTab"
  | "settings";

export type ShortcutMap = Record<ShortcutId, string>;

/** 应用设置 */
export interface Settings {
  theme: ThemeMode;
  accent: AccentName;
  fontSize: number;
  fontFamily: string;
  tabSize: number;
  wordWrap: boolean;
  showLineNumbers: boolean;
  autoSave: boolean;
  autoSaveInterval: number;
  recentFilesLimit: number;
  shortcuts: ShortcutMap;
  autoCheckUpdate: boolean;
  /** 大纲显示的最大标题等级（1-6） */
  outlineMaxLevel: number;
  /** 红绿色弱友好模式（采用 Okabe-Ito 无障碍配色与非纯色视觉标识） */
  colorblindMode: boolean;
}

/** 未保存变更弹窗结果 */
export type UnsavedChoice = "save" | "discard" | "cancel";

/** 外部修改冲突弹窗结果 */
export type ConflictChoice = "local" | "external" | "saveas";

export interface ConfirmRequest {
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  danger?: boolean;
}
