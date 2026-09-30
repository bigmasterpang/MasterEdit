//! DOCX 只读查看：解密 → 解包 → OOXML 解析成「块模型」。
//!
//! 设计要点（详见 `docs/plan-docx.md`）：
//! 1. **只读**：不写回文件，因此没有"改写 OOXML 弄坏文档"的风险；也不产生临时文件；
//! 2. **企业加密透明**：复用 `office::read_raw_and_plain`（亿赛通等在内存里解密，
//!    **绝不落明文临时文件**）；
//! 3. **不认识的节点不丢**：解析成块模型，无法呈现的对象（SmartArt / 图表 / OLE /
//!    文本框 / OMML 公式 / altChunk）保留为「占位块 + 中文说明」，而不是静默跳过；
//! 4. **绝不 panic**：XML 损坏、部件缺失、关系缺失、索引越界一律降级处理（`Option`/`Result` 兜底）；
//! 5. 解析与渲染分离：这里只产出结构化数据，前端负责排版（结构忠实，不追求像素级还原）。
//!
//! 样式继承链（决定"像不像"的就是它）。两条链都按"从最内层往外找"的顺序解析，
//! 每个属性取第一个明确写了值的层：
//! ```text
//! 字符属性（字体/字号/颜色/粗斜体/下划线）：run 的 rPr → 字符样式(含 basedOn)
//!     → 段落标记的 rPr(pPr/rPr) → 段落样式(含 basedOn 逐级) → docDefaults/rPr
//! 段落属性（对齐/缩进/间距/大纲级别）：直接 pPr → 编号级别自带的缩进
//!     → 段落样式(含 basedOn 逐级) → docDefaults/pPr
//! ```
//!
//! 真实样本（WPS 导出的周报）里 `docDefaults` 只写了 **主题字体**（`w:asciiTheme="minorHAnsi"`），
//! 一个具体字体名都没有 —— 所以必须连 `word/theme/theme1.xml` 的 `a:fontScheme` 一起解析，
//! 否则整篇文档的字体都是空的（"等线"这个词在这份文档的 document.xml / styles.xml 里
//! 一次都没出现，只在 theme1.xml 里有）。
//!
//! **已知保真度缺口**（有意不做，详见 `docs/plan-docx.md` 的"已知取舍"）：
//! - `w:contextualSpacing`（同样式段落之间不加空）未实现 —— 它是**渲染规则**而不是数值，
//!   要让前端遵守就得新增契约字段（`Block`/`ParagraphBlock` 加标志），收益相对有限，
//!   所以本版不导出；段落间距按 XML 里的实际数值渲染。
//! - 页眉页脚只取**默认**那一套（`w:headerReference`/`w:footerReference` 的 `w:type="default"`，
//!   没有 default 时取第一个存在的）：**"首页不同"（`w:titlePg`）与奇偶页不区分**，
//!   所以首页也按默认页眉页脚显示。多节的页眉页脚同理只取最后一节（`w:body/w:sectPr`）。
//! - 浮动图片/文本框按内联显示、不做文字环绕排版（`wrap` 字段只用来判断"是不是浮动对象"）。
//! - SmartArt / 图表 / OLE / OMML 公式只给占位块。

use std::collections::HashMap;
use std::io::{Cursor, Read};
use std::path::Path;
use std::sync::{Arc, Mutex, OnceLock};

use serde::Serialize;
use zip::ZipArchive;

use super::{file, office};

/* ================================================================================== */
/* 一、块模型：序列化给前端的数据结构（前端按 kind 分派渲染）                          */
/* ================================================================================== */

/// 文档中的一个有序块。JSON 里带 `kind` 判别字段（`paragraph` / `table` / `image` /
/// `shape` / `textBox` / `pageBreak` / `unsupported`），字段名统一 camelCase，
/// 前端可以直接写成 TS 联合类型。
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum Block {
    /// 普通段落（含表格单元格里的段落）
    Paragraph(ParagraphBlock),
    /// 表格（行 × 单元格，单元格里是块列表，可含嵌套表格）
    Table(TableBlock),
    /// 独立成块的图片：段落里没有文字、只有图片时，一张图一个块
    Image(ImageBlock),
    /// 简单图形：合同/表单里用「直线 / 矩形」画的横线与方框
    /// （VML 的 `v:line` / `v:rect` / `v:roundrect` / `v:oval` / `v:hr`，
    ///   DrawingML 的 `wps:wsp` + `a:prstGeom`）。识别不出来的形状仍走 [`Block::Unsupported`]。
    Shape(ShapeBlock),
    /// 文本框：里面的内容**结构化**成块（封面、表单、页脚里的框都用它）。
    /// 空文本框、或压根没有内容的形状才退回 [`Block::Unsupported`]。
    TextBox(TextBoxBlock),
    /// 显式分页符（`<w:br w:type="page"/>` 独占一段时）
    PageBreak,
    /// 无法呈现的对象 —— **绝不静默丢失**，前端显示占位卡片 + 中文说明
    Unsupported { label: String, detail: String },
}

/// 段落块。所有长度单位已换算成 **pt / px**（XML 里的 twip、半磅、EMU 都已折算）。
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphBlock {
    /// 段落样式名（`w:pStyle` 指向的样式在 `w:name` 里显示的名字，如「heading 4」）
    pub style: Option<String>,
    /// 段落样式 id（`w:pStyle@w:val`，WPS 里常是 `4`/`a9` 这种）
    pub style_id: Option<String>,
    /// 段落内的文本片段
    pub runs: Vec<Run>,
    /// 段落纯文本（各 run 文本拼接，`\n`=软换行、`\t`=制表符）；查找与"复制全部文本"直接用它
    pub text: String,
    /// 对齐：left / center / right / both（两端对齐）/ distribute
    pub align: Option<String>,
    /// 左缩进（pt）
    pub indent_left_pt: Option<f32>,
    /// 右缩进（pt）
    pub indent_right_pt: Option<f32>,
    /// 首行缩进（pt，**负值表示悬挂缩进**，列表常用）
    pub indent_first_line_pt: Option<f32>,
    /// 段前间距（pt）
    pub space_before_pt: Option<f32>,
    /// 段后间距（pt）
    pub space_after_pt: Option<f32>,
    /// 行距
    pub line_spacing: Option<LineSpacing>,
    /// 列表（项目符号 / 编号）信息，含该段应显示的前缀文本
    pub list: Option<ListInfo>,
    /// 大纲级别 0..8（标题样式 → 前端大纲侧栏用）
    pub outline_level: Option<u8>,
    /// `w:pageBreakBefore`：段前分页
    pub page_break_before: bool,
    /// 段落内部含分页符（与文字混排时；独占一段时整个块是 [`Block::PageBreak`]）
    pub page_break: bool,
    /// 分节符类型（nextPage / continuous / evenPage / oddPage）；有值时前端可画一条淡色"分页提示线"
    pub section_break: Option<String>,
    /// 段落边框（`w:pPr/w:pBdr`）：四边都为空时整个字段是 `null`。
    /// 合同/表单里大量用"空段落 + 下边框"画横线，这是除图形外最常见的画线方式。
    pub borders: Option<ParagraphBorders>,
}

/// 段落边框。只包含 `w:val` 不为 `none`/`nil` 的边（明确不要边框的边不出现）。
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphBorders {
    pub top: Option<BorderSpec>,
    pub bottom: Option<BorderSpec>,
    pub left: Option<BorderSpec>,
    pub right: Option<BorderSpec>,
}

/// 一条边框（`w:top` / `w:bottom` / …）
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BorderSpec {
    /// `w:val`：single / double / dashed / dotted / thick / wave …（`none`/`nil` 不会出现在这里）
    pub style: String,
    /// 线宽（pt）；`w:sz` 的单位是**八分之一磅**，所以 ÷8
    pub width_pt: f64,
    /// 颜色 RRGGBB（`w:color="auto"` → `null`，前端给默认色）
    pub color: Option<String>,
    /// 边框与文字的间距（pt，`w:space`；文档里没写就是 `null`）
    pub space_pt: Option<f64>,
}

/// 行距。`multiple` 时 value 是倍数（1.0 = 单倍），`exact`/`atLeast` 时 value 是 pt。
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct LineSpacing {
    pub kind: String,
    pub value: f32,
}

/// 列表信息：前缀文本已经算好（含多级编号的父级计数）
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ListInfo {
    /// `w:numId`
    pub num_id: u32,
    /// `w:ilvl`（0 起）
    pub level: u8,
    /// 有序（编号 1. /（一）/ 1.1）还是无序（项目符号 •）
    pub ordered: bool,
    /// 该段应显示的前缀文本，例如 `•` / `1.` / `1.1` / `（一）`；项目符号级就是那个符号
    pub prefix: String,
    /// `w:numFmt` 原值（decimal / bullet / chineseCounting …；未知时原样给出）
    pub format: String,
    /// 标记与正文之间的间隔方式：tab / space / nothing（`w:suff`）
    pub suffix: String,
}

/// 会跟着分页变的域类型。
///
/// 这类域在 XML 里存的是**上次保存时的缓存结果**（页脚里的 `2`、`30`、`I`、`II` 都是它），
/// 前端拿到实际页序后要用实时值替换，所以标记要一路带到 run 上（`Run::field`）。
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "UPPERCASE")]
pub enum FieldKind {
    /// `PAGE`：当前页码
    Page,
    /// `NUMPAGES`：总页数
    NumPages,
}

/// 文本片段（run）。布尔值都是"最终生效"的结果，不是原始值。
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Run {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    /// 西文字体（`w:rFonts/@w:ascii`→`@w:hAnsi`，含主题字体解析）
    pub font: Option<String>,
    /// 中日韩字体（`w:rFonts/@w:eastAsia`，含主题字体解析）；中文按它渲染
    pub font_east_asia: Option<String>,
    /// 字号（pt；文档里是半磅，故 ÷2）
    pub size_pt: Option<f32>,
    /// 文字颜色（RRGGBB，`w:color`；`auto` 视为未指定）
    pub color: Option<String>,
    /// 高亮或底纹（RRGGBB，`w:highlight` → `w:shd@w:fill`）
    pub highlight: Option<String>,
    /// superscript / subscript（`w:vertAlign`）
    pub vert_align: Option<String>,
    /// 字符样式 id（`w:rStyle@w:val`）
    pub style_id: Option<String>,
    /// 这个 run 的文本是**域缓存结果**（`PAGE` / `NUMPAGES`）时要替换成实时值；
    /// 其余 run 一律 `None`
    pub field: Option<FieldKind>,
}

/// 表格块
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableBlock {
    pub rows: Vec<TableRow>,
    /// 列宽（px，来自 `w:tblGrid`；文档里没写就是空数组）
    pub columns: Vec<f32>,
    /// 表格总宽（px，仅当 `w:tblW` 用 dxa/twip 给出时才有）
    pub width_px: Option<f32>,
    /// 表格对齐：left / center / right
    pub align: Option<String>,
    /// 表格是否有可见边框（表格边框 + 任一单元格边框，两者取或）
    pub borders: bool,
    /// 表格样式 id（`w:tblStyle@w:val`）
    pub style_id: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    /// 行高（px，`w:trHeight`；`hRule="auto"` 时它是最小高度）
    pub height_px: Option<f32>,
    /// `w:tblHeader`：跨页重复的表头行
    pub header: bool,
}

/// 表格单元格。合并信息按 OOXML 原样给出：
/// - `grid_span > 1`：横向合并了这么多列
/// - `v_merge = restart`：纵向合并的**起点**；后续同列的 `continue` 表示"和上面是同一格"
///   （前端据此算 rowspan，continue 单元格一般不重复渲染内容）
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    /// 单元格内的块（多个段落，甚至嵌套表格）
    pub blocks: Vec<Block>,
    /// 单元格纯文本（各段用 `\n` 连接）；「复制表格为制表符文本」直接用
    pub text: String,
    /// 横向合并列数（`w:gridSpan`，至少 1）
    pub grid_span: usize,
    pub v_merge: VMerge,
    /// 单元格宽度（px，仅当 `w:tcW` 用 dxa 给出时）
    pub width_px: Option<f32>,
    /// 底纹（RRGGBB，`w:shd@w:fill`）
    pub shading: Option<String>,
    /// 垂直对齐：top / center / bottom
    pub v_align: Option<String>,
    /// 单元格四条边是否有可见边框（已把表格级 `w:tblBorders` 作为兜底算进来）
    pub borders: CellBorders,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum VMerge {
    #[default]
    None,
    /// 纵向合并的起点（`<w:vMerge w:val="restart"/>`）
    Restart,
    /// 纵向合并的延续（`<w:vMerge/>` 或 `w:val="continue"`）
    Continue,
}

#[derive(Debug, Clone, Copy, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellBorders {
    pub top: bool,
    pub left: bool,
    pub bottom: bool,
    pub right: bool,
}

/// 图片块（`word/media/` 里的媒体文件；前端走已有 asset 协议按需取）
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageBlock {
    /// 包内路径，如 `word/media/image1.png`
    pub media: String,
    /// `wp:docPr@name`（原始插入名，Word 里是「图片 1」）
    pub name: Option<String>,
    /// 替代文本（`wp:docPr@descr`）
    pub alt: Option<String>,
    /// 显示宽（px，EMU ÷ 9525）
    pub width_px: f32,
    /// 显示高（px）
    pub height_px: f32,
}

/// 简单图形（合同/表单里的横线、方框）。坐标与尺寸都用 **pt**。
///
/// 注意 `xPt` / `yPt` 的含义：Word 对内联（`wp:inline`）图形**不记录**它在段落里的水平
/// 偏移（那是排版算出来的），所以这里给的是形状**自身声明**的偏移（DrawingML 的 `a:off`、
/// VML style 里的 `left`/`margin-left`），拿不到就是 `0` —— 前端按"就在这个 run 的位置"处理。
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeBlock {
    /// 形状种类：line / rect / roundRect / ellipse
    pub shape: String,
    /// 相对段落内容区左上角的横向偏移（pt，拿不到就是 0）
    pub x_pt: f64,
    /// 纵向偏移（pt，同样只在形状自己声明了才有意义）
    pub y_pt: f64,
    /// 宽（pt）
    pub width_pt: f64,
    /// 高（pt）；直线的高度为 0 时按线宽（拿不到就 1pt）兜底，前端好定位
    pub height_pt: f64,
    /// 线宽（pt）；`a:ln@w`（EMU ÷ 12700）或 `strokeweight`
    pub line_width_pt: Option<f64>,
    /// 线色 RRGGBB（拿不到给 `null`，前端用默认黑）
    pub line_color: Option<String>,
    /// 填充色 RRGGBB（无填充 / 拿不到都是 `null`）
    pub fill_color: Option<String>,
    /// 虚线样式（DrawingML 的 `a:prstDash` 原值，或 VML `dashstyle` 归一化后的名字）
    pub dash: Option<String>,
    /// 是否竖线（仅 `line` 有意义：高 > 宽）
    pub vertical: bool,
}

/// 文本框。**内容是真解析出来的块**（段落/表格/图片/嵌套块走同一套样式层叠与编号），
/// 不是一张占位卡片 —— 企业标准的封面整页、表单里的签字框都是文本框。
///
/// `xPt`/`yPt` 与 [`ShapeBlock`] 同义：Word 对**内联**文本框不记录段落内偏移，
/// 拿不到就是 0（前端按"就在这个 run 的位置"处理）；浮动文本框给的是它自己声明的偏移。
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBoxBlock {
    /// 文本框内部的块
    pub blocks: Vec<Block>,
    /// 相对段落内容区左上角的横向偏移（pt；拿不到为 0）
    pub x_pt: f64,
    /// 纵向偏移（pt）
    pub y_pt: f64,
    /// 宽（pt；拿不到给 200pt 兜底）
    pub width_pt: f64,
    /// 高（pt；拿不到给 40pt 兜底）
    pub height_pt: f64,
    /// 填充色 RRGGBB（无填充 / 拿不到为 `null`）
    pub fill_color: Option<String>,
    /// 边框色 RRGGBB
    pub border_color: Option<String>,
    /// 边框线宽（pt）
    pub border_width_pt: Option<f64>,
    /// 环绕方式原值：square / none / tight / through / topAndBottom
    /// （前端只用它判断"是不是浮动对象"；内联文本框给 `none`）
    pub wrap: String,
}

/// `document_blocks` 的返回：窗口化的块 + 总数（前端虚拟滚动用）
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BlockPage {
    /// 文档里的块总数（**含表格单元格内的块**不参与，这里只算顶层块）
    pub total: usize,
    /// 实际返回的起始下标（已夹到 [0, total]）
    pub from: usize,
    pub blocks: Vec<Block>,
    /// 是否为亿赛通等透明加密文档（已在内存中解密）
    pub encrypted: bool,
    /// 页面几何（"一页一页"显示时前端算分页要用）；文档里没有 `w:sectPr` 时为 `None`
    pub page: Option<PageGeometry>,
    /// **文档自己的页眉**（`w:headerReference`，默认页眉）解析出的块；没有就是 `None`
    pub header: Option<Vec<Block>>,
    /// **文档自己的页脚**（`w:footerReference`，默认页脚）解析出的块
    pub footer: Option<Vec<Block>>,
}

/// 页面几何：取自 `w:sectPr` 的 `w:pgSz` / `w:pgMar`（单位统一 pt，twip ÷ 20）。
///
/// 只描述**纸张与页边距**，不承诺页码/分节位置 —— 分页由前端按内容高度切。
/// 单个属性缺失或畸形（`w:w="0"`、非数字）时按 A4 兜底，不 panic。
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageGeometry {
    /// 纸张宽（pt，`w:pgSz/@w:w`）
    pub width_pt: f64,
    /// 纸张高（pt，`w:pgSz/@w:h`）
    pub height_pt: f64,
    /// 上边距（pt，`w:pgMar/@w:top`；**可为负**，Word 允许）
    pub margin_top_pt: f64,
    /// 右边距（pt）
    pub margin_right_pt: f64,
    /// 下边距（pt）
    pub margin_bottom_pt: f64,
    /// 左边距（pt）
    pub margin_left_pt: f64,
    /// `w:pgSz/@w:orient == "landscape"`（注意：Word 写横版时会把 w/h 直接写成横向尺寸）
    pub landscape: bool,
}

/// 一处查找命中（块级：一个块最多一条，`block` 是顶层块下标，供前端滚动定位）
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FindHit {
    /// 顶层块下标（可直接喂给 `document_blocks(from = block)`）
    pub block: usize,
    /// 命中处的上下文片段（首尾可能带 `…`）
    pub text: String,
}

/// 打开文档时读到的包信息（供状态栏与调试用）
#[derive(Debug, Serialize)]
pub struct DocumentInfo {
    /// 是否为亿赛通等透明加密文档（已在内存中解密）
    pub encrypted: bool,
    /// 文档是否可编辑 —— DOCX 目前**一律只读**（编辑交给 Word / WPS，见「打开方式」）
    pub editable: bool,
    /// 包内部件（名称 + 原始大小），按大小降序
    pub parts: Vec<DocumentPart>,
    /// 段落数（初步统计，用于状态栏显示规模）
    pub paragraphs: usize,
    /// 表格数
    pub tables: usize,
    /// 图片数（`word/media/` 下的文件数）
    pub images: usize,
}

#[derive(Debug, Serialize)]
pub struct DocumentPart {
    pub name: String,
    pub size: u64,
}

/* ================================================================================== */
/* 二、极简 XML DOM                                                                   */
/* ================================================================================== */

/// 一个极简 DOM。docx 的 XML 单部件通常几百 KB ~ 几 MB，建成树再按名字查最省心，
/// 也避免手写状态机在"未知节点"上丢内容。用完即随 `ParsedDocument` 一起释放。
/// （`Clone` 是给页眉/页脚的"独立故事"用的：那些部件需要自己持有一棵树。）
#[derive(Clone)]
struct XmlNode {
    /// **原始限定名**（如 `w:pStyle`），保留下来是为了在占位说明里报出真实节点名
    name: String,
    attrs: Vec<(String, String)>,
    children: Vec<XmlChild>,
}

#[derive(Clone)]
enum XmlChild {
    Elem(XmlNode),
    Text(String),
}

/// 取限定名的本地名（`w:pStyle` → `pStyle`）。
/// 全程按本地名匹配：Word/WPS/严格 ISO 文档的前缀各不相同，按前缀匹配会漏。
fn local_name(qname: &str) -> &str {
    match qname.rsplit_once(':') {
        Some((_, local)) => local,
        None => qname,
    }
}

impl XmlNode {
    fn local(&self) -> &str {
        local_name(&self.name)
    }

    /// 按**限定名**精确取属性（部件内部无前缀的属性用它，如 `Id` / `Target`）
    fn attr(&self, name: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.as_str())
    }

    /// 按**本地名**取属性（`w:val` / `r:embed` / `w:w` 都用它）
    fn attr_local(&self, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(key, _)| local_name(key) == local)
            .map(|(_, value)| value.as_str())
    }

    fn children(&self) -> impl Iterator<Item = &XmlNode> {
        self.children.iter().filter_map(|child| match child {
            XmlChild::Elem(node) => Some(node),
            XmlChild::Text(_) => None,
        })
    }

    /// 直接子元素里第一个本地名匹配的
    fn child(&self, local: &str) -> Option<&XmlNode> {
        self.children().find(|node| node.local() == local)
    }

    /// 全部直接子元素里本地名匹配的
    fn children_named<'a>(&'a self, local: &'a str) -> impl Iterator<Item = &'a XmlNode> + 'a {
        self.children().filter(move |node| node.local() == local)
    }

    /// 直接文本子节点拼接（`<w:t>文字</w:t>` 用它取文字）
    fn texts(&self) -> String {
        let mut out = String::new();
        for child in &self.children {
            if let XmlChild::Text(text) = child {
                out.push_str(text);
            }
        }
        out
    }

    /// 深度优先找第一个本地名匹配的后代（**迭代式**，避免深文档递归爆栈）
    fn find_descendant(&self, local: &str) -> Option<&XmlNode> {
        let mut stack: Vec<&XmlNode> = self.children().collect();
        stack.reverse();
        while let Some(node) = stack.pop() {
            if node.local() == local {
                return Some(node);
            }
            let mut next: Vec<&XmlNode> = node.children().collect();
            next.reverse();
            stack.extend(next);
        }
        None
    }

    /// 后代里是否存在某本地名的节点
    fn has_descendant(&self, local: &str) -> bool {
        self.find_descendant(local).is_some()
    }

    /// 后代里所有 `w:t` / `m:t` 的文字（占位说明里回显"丢掉了什么内容"）。
    /// 用显式栈按**文档顺序**遍历（递归写法遇到深文档可能爆栈）
    fn collect_text(&self) -> String {
        let mut out = String::new();
        let mut stack: Vec<&XmlNode> = vec![self];
        while let Some(node) = stack.pop() {
            if node.local() == "t" {
                out.push_str(&node.texts());
                continue;
            }
            // 逆序压栈 → 弹出顺序与文档顺序一致
            let children: Vec<&XmlNode> = node.children().collect();
            for child in children.into_iter().rev() {
                stack.push(child);
            }
        }
        out
    }
}

/// 解析 XML 文本。**损坏的 XML 不返回 Err 而是尽力而为**：能读多少读多少，
/// 结束标签不匹配、裸 `&` 都容忍（只读查看场景下"少显示一点"远好过"整篇打不开"）。
fn parse_xml(text: &str) -> Option<XmlNode> {
    use quick_xml::events::Event;

    if text.trim().is_empty() {
        return None;
    }
    let mut reader = quick_xml::reader::Reader::from_str(text);
    {
        let config = reader.config_mut();
        config.check_end_names = false; // 结束标签名不匹配也继续
        config.allow_unmatched_ends = true; // 多出来的结束标签直接忽略
        config.allow_dangling_amp = true; // 裸 & 不报错
        config.trim_text(false); // 保留 <w:t xml:space="preserve"> 里的空白
    }

    let mut roots: Vec<XmlNode> = Vec::new();
    let mut stack: Vec<XmlNode> = Vec::new();

    loop {
        let event = match reader.read_event() {
            Ok(event) => event,
            // 解析不动了就收手：已经读到的部分照常用，绝不 panic
            Err(_) => break,
        };
        match event {
            Event::Start(start) => stack.push(new_node(&start)),
            Event::Empty(start) => {
                let node = new_node(&start);
                attach(&mut roots, &mut stack, node);
            }
            Event::End(_) => {
                if let Some(node) = stack.pop() {
                    attach(&mut roots, &mut stack, node);
                }
            }
            Event::Text(text) => push_text(&mut stack, text.into_inner().as_ref()),
            Event::CData(data) => push_text(&mut stack, data.into_inner().as_ref()),
            // quick-xml 0.42 把实体引用单独作为事件抛出（`&amp;` / `&#x4E2D;`）
            Event::GeneralRef(reference) => {
                push_text(&mut stack, &resolve_entity(reference.into_inner().as_ref()))
            }
            Event::Eof => break,
            _ => {}
        }
    }
    while let Some(node) = stack.pop() {
        attach(&mut roots, &mut stack, node);
    }

    match roots.len() {
        0 => None,
        1 => roots.pop(),
        // 多个根（严重损坏）：合成一个容器，后续按名字仍能找到 w:body
        _ => Some(XmlNode {
            name: "#fragment".to_string(),
            attrs: Vec::new(),
            children: roots.into_iter().map(XmlChild::Elem).collect(),
        }),
    }
}

/// 从开始标签建节点（属性值做实体反转义；反转义失败就退回原始值，不丢数据）
fn new_node(start: &quick_xml::events::BytesStart<'_>) -> XmlNode {
    let mut attrs = Vec::with_capacity(4);
    for attr in start.attributes().flatten() {
        let key = attr.key.as_ref().to_string();
        // normalized_value：属性值按 XML 规范归一化（换行/制表变空格 + 预定义实体反转义）
        let value = match attr.normalized_value(quick_xml::XmlVersion::Implicit1_0) {
            Ok(value) => value.into_owned(),
            Err(_) => attr.value.as_ref().to_string(),
        };
        attrs.push((key, value));
    }
    XmlNode {
        name: start.name().as_ref().to_string(),
        attrs,
        children: Vec::new(),
    }
}

/// 把节点挂到栈顶（或作为根）
fn attach(roots: &mut Vec<XmlNode>, stack: &mut [XmlNode], node: XmlNode) {
    match stack.last_mut() {
        Some(parent) => parent.children.push(XmlChild::Elem(node)),
        None => roots.push(node),
    }
}

fn push_text(stack: &mut [XmlNode], text: &str) {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlChild::Text(text.to_string()));
    }
}

/// 实体引用 → 文本（`amp` / `#x4E2D` / …；认不出来就原样保留，不丢字符）
fn resolve_entity(name: &str) -> String {
    if let Some(rest) = name.strip_prefix('#') {
        let code = match rest.strip_prefix(['x', 'X']) {
            Some(hex) => u32::from_str_radix(hex, 16).ok(),
            None => rest.parse::<u32>().ok(),
        };
        return code
            .and_then(char::from_u32)
            .map(|ch| ch.to_string())
            .unwrap_or_default();
    }
    match name {
        "amp" => "&".to_string(),
        "lt" => "<".to_string(),
        "gt" => ">".to_string(),
        "quot" => "\"".to_string(),
        "apos" => "'".to_string(),
        other => format!("&{other};"),
    }
}

/* ================================================================================== */
/* 三、样式继承链（docDefaults → 段落样式 basedOn → 字符样式 → 直接格式）              */
/* ================================================================================== */

/// 字符级属性。全部是 `Option`：`Some` = 这条链上明确写了值，`None` = 没写（继续往下找）。
/// 颜色/字体/字号都是"最靠内层说了算"，所以用 Option 而不是带默认值的 bool。
#[derive(Debug, Clone, Default)]
struct RunProps {
    font_latin: Option<String>,
    font_ea: Option<String>,
    size_half_pt: Option<f32>,
    color: Option<String>,
    highlight: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    vert_align: Option<String>,
}

impl RunProps {
    /// 用 `lower`（优先级更低的一层）补齐自己还没有的字段。
    /// 层叠时从高优先级往低优先级依次调用，每个字段自然取到最靠内层的值。
    fn apply_over(&mut self, lower: &RunProps) {
        fn fill<T: Clone>(slot: &mut Option<T>, lower: &Option<T>) {
            if slot.is_none() {
                *slot = lower.clone();
            }
        }
        fill(&mut self.font_latin, &lower.font_latin);
        fill(&mut self.font_ea, &lower.font_ea);
        fill(&mut self.size_half_pt, &lower.size_half_pt);
        fill(&mut self.color, &lower.color);
        fill(&mut self.highlight, &lower.highlight);
        fill(&mut self.bold, &lower.bold);
        fill(&mut self.italic, &lower.italic);
        fill(&mut self.underline, &lower.underline);
        fill(&mut self.strike, &lower.strike);
        fill(&mut self.vert_align, &lower.vert_align);
    }
}

/// 段落级属性（长度统一换算成 pt）
///
/// 有三类设置**不能**在解析 `w:pPr` 时就算完，必须等整条层叠链走完：
/// - **字符单位**的缩进（`w:firstLineChars` / `w:leftChars` / `w:hangingChars`）：
///   1 字符 = 本段**最终生效的字号**，字号要等样式链算完才知道；
/// - **行单位**的段间距（`w:beforeLines` / `w:afterLines`）：1 行 = 本段最终行距折出的单行高度；
/// - **自动段间距**（`w:beforeAutospacing` / `w:afterAutospacing`）：开了「自动」时段前/段后
///   要盖掉样式与 docDefaults 里那些常常是 0 的具体值。
///
/// 所以这三类先用 `Option` 原样带过来，层叠结束后由 [`ParaProps::resolve_char_indents`] /
/// [`ParaProps::resolve_spacing`] 定稿。
#[derive(Debug, Clone, Default)]
struct ParaProps {
    align: Option<String>,
    indent_left_pt: Option<f32>,
    indent_right_pt: Option<f32>,
    first_line_pt: Option<f32>,
    space_before_pt: Option<f32>,
    space_after_pt: Option<f32>,
    line: Option<LineSpacing>,
    outline_level: Option<u8>,
    /// `w:leftChars` / `w:startChars`（单位：百分之一字符，200 = 2 字符）
    left_chars: Option<f32>,
    /// `w:rightChars` / `w:endChars`
    right_chars: Option<f32>,
    /// `w:firstLineChars`（正）或 `-w:hangingChars`（负）
    first_line_chars: Option<f32>,
    /// `w:beforeLines`（单位：百分之一行，200 = 2 行）
    before_lines: Option<f32>,
    /// `w:afterLines`
    after_lines: Option<f32>,
    /// `w:beforeAutospacing`（自动段前间距）
    auto_before: Option<bool>,
    /// `w:afterAutospacing`（自动段后间距）
    auto_after: Option<bool>,
    /// `w:pBdr` 段落边框（四边都空时解析成 None）
    borders: Option<ParagraphBorders>,
}

impl ParaProps {
    fn apply_over(&mut self, lower: &ParaProps) {
        fn fill<T: Clone>(slot: &mut Option<T>, lower: &Option<T>) {
            if slot.is_none() {
                *slot = lower.clone();
            }
        }
        /// 「同一个属性的两种写法」的层叠：**层优先，同层内 twips 优先**。
        ///
        /// `w:ind` 的 twips 与 `Chars`、`w:spacing` 的 twips 与 `Lines` 都是**同一个属性的
        /// 两种编码**（Word/WPS 写文档时会在同一层把两种一起写下，这时用 twips —— 那是它
        /// 排版用的值）。但**跨层**必须让"更靠内层的写法整组胜出"：OOXML 里直接格式本来就
        /// 优先于样式，与用哪种单位写无关 —— 所以这里按"组"来填，而不是逐字段填。
        fn fill_pair(
            own_twips: &mut Option<f32>,
            own_units: &mut Option<f32>,
            lower_twips: &Option<f32>,
            lower_units: &Option<f32>,
        ) {
            if own_twips.is_none() && own_units.is_none() {
                if lower_twips.is_some() {
                    *own_twips = *lower_twips;
                } else {
                    *own_units = *lower_units;
                }
            }
        }
        fill(&mut self.align, &lower.align);
        fill(&mut self.line, &lower.line);
        fill(&mut self.outline_level, &lower.outline_level);
        fill(&mut self.auto_before, &lower.auto_before);
        fill(&mut self.auto_after, &lower.auto_after);
        fill(&mut self.borders, &lower.borders);
        // 缩进：twips 与 Chars 是一组
        fill_pair(
            &mut self.indent_left_pt,
            &mut self.left_chars,
            &lower.indent_left_pt,
            &lower.left_chars,
        );
        fill_pair(
            &mut self.indent_right_pt,
            &mut self.right_chars,
            &lower.indent_right_pt,
            &lower.right_chars,
        );
        fill_pair(
            &mut self.first_line_pt,
            &mut self.first_line_chars,
            &lower.first_line_pt,
            &lower.first_line_chars,
        );
        // 段间距：twips 与 Lines 是一组
        fill_pair(
            &mut self.space_before_pt,
            &mut self.before_lines,
            &lower.space_before_pt,
            &lower.before_lines,
        );
        fill_pair(
            &mut self.space_after_pt,
            &mut self.after_lines,
            &lower.space_after_pt,
            &lower.after_lines,
        );
    }

    /// 把「字符单位」的缩进折算成 pt（用该段**最终生效的字号**）。
    ///
    /// 优先级（从高到低，与 [`ParaProps::resolve_spacing`] 同一套语义）：
    /// 1. 段落自己写的 twips（`w:left` / `w:firstLine` / `w:hanging`）
    /// 2. 段落自己写的 Chars（`w:leftChars` / `w:firstLineChars` / `w:hangingChars`）
    /// 3. 层叠结果（编号级别 → 段落样式 → docDefaults）
    ///
    /// 前两级在 [`ParaProps::apply_over`] 里就已经分好（组内 twips 优先、跨层按层），
    /// 走到这里每一组最多只剩一个值，所以只需把 Chars 折算出来。
    /// 为什么必须等层叠结束：Word / WPS 里「首行缩进 2 字符」是相对本段字号的 ——
    /// 同一份 XML 在 10.5pt 段落里是 21pt、在 24pt 段落里是 48pt。
    fn resolve_char_indents(&mut self, font_size_pt: f32) {
        let per_char = |chars: f32| chars / 100.0 * font_size_pt;
        if self.indent_left_pt.is_none() {
            if let Some(chars) = self.left_chars {
                self.indent_left_pt = Some(per_char(chars));
            }
        }
        if self.indent_right_pt.is_none() {
            if let Some(chars) = self.right_chars {
                self.indent_right_pt = Some(per_char(chars));
            }
        }
        if self.first_line_pt.is_none() {
            if let Some(chars) = self.first_line_chars {
                self.first_line_pt = Some(per_char(chars));
            }
        }
    }

    /// 段前/段后定稿。优先级（从高到低，与 [`ParaProps::resolve_char_indents`] 同一套语义）：
    /// 1. 段落自己写的 twips（`w:before` / `w:after`）
    /// 2. 段落自己写的行单位（`w:beforeLines` / `w:afterLines`）→ 百分之一行 × 行高基准
    /// 3. 自动段间距（`w:beforeAutospacing` / `w:afterAutospacing` → [`AUTO_SPACE_PT`]）
    /// 4. 层叠结果（编号级别 → 段落样式 → docDefaults）
    ///
    /// 前两级在 [`ParaProps::apply_over`] 里已经按"层优先、同层内 twips 优先"分好；
    /// 这里只处理 2 的折算，以及 **Auto 为什么能压过第 4 级**：现实文档的 docDefaults 常写
    /// `w:before="0" w:after="0"`，照搬就会比 Word 紧一大截 —— Word 里勾了「自动」，
    /// 段落对话框显示的是 Auto，样式里的具体值并不参与。段落自己写了值（第 1、2 级）时
    /// 才让位（`explicit_before` / `explicit_after` 就是"段落自己写死了 twips"）。
    ///
    /// 为什么行单位也要等层叠结束：`beforeLines` 是「几行」，1 行 = 本段最终行距折出的
    /// 单行高度（字号与行距都要等样式链算完）。
    fn resolve_spacing(&mut self, font_size_pt: f32, explicit_before: bool, explicit_after: bool) {
        let line_height_pt = self.line_height_pt(font_size_pt);
        if !explicit_before {
            self.space_before_pt = self
                .before_lines
                .map(|lines| lines / 100.0 * line_height_pt)
                .or_else(|| (self.auto_before == Some(true)).then_some(AUTO_SPACE_PT))
                .or(self.space_before_pt);
        }
        if !explicit_after {
            self.space_after_pt = self
                .after_lines
                .map(|lines| lines / 100.0 * line_height_pt)
                .or_else(|| (self.auto_after == Some(true)).then_some(AUTO_SPACE_PT))
                .or(self.space_after_pt);
        }
    }

    /// 「行高基准」：`beforeLines` / `afterLines` 折算成 pt 时用的单行高度。
    /// - `exact` / `atLeast` → 行距自己的 pt 值；
    /// - `multiple` → 字号 × [`SINGLE_LINE_FACTOR`] × 倍数；
    /// - 没写行距 → 字号 × [`SINGLE_LINE_FACTOR`]（单倍行距）。
    fn line_height_pt(&self, font_size_pt: f32) -> f32 {
        let single = font_size_pt * SINGLE_LINE_FACTOR;
        let height = match &self.line {
            Some(spacing) if spacing.kind != "multiple" => spacing.value,
            Some(spacing) => single * spacing.value,
            None => single,
        };
        // 损坏/异常的 XML 可能给出 0 或负的行距，别让它把段间距也变成 0
        if height > 0.0 {
            height
        } else {
            single
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StyleKind {
    Paragraph,
    Character,
    Other,
}

#[derive(Debug, Clone)]
struct StyleDef {
    /// `w:name`（Word 界面里显示的名字，如 `heading 4`）
    name: Option<String>,
    based_on: Option<String>,
    rpr: RunProps,
    ppr: ParaProps,
}

#[derive(Debug, Default)]
struct StyleSheet {
    /// `w:docDefaults/w:rPrDefault/w:rPr`
    doc_rpr: RunProps,
    /// `w:docDefaults/w:pPrDefault/w:pPr`
    doc_ppr: ParaProps,
    styles: HashMap<String, StyleDef>,
    /// `w:default="1"` 的段落样式（通常是 Normal）：段落没写 `w:pStyle` 时用它
    default_para_style: Option<String>,
    /// `w:default="1"` 的字符样式（通常是 Default Paragraph Font）
    default_char_style: Option<String>,
}

/// 样式继承最大层数：正常文档 3~5 层，给到 10 足够；
/// 同时用 visited 集合防止 `basedOn` 成环（WPS/手工改过的文档真出现过）。
const MAX_STYLE_DEPTH: usize = 10;

/// 「自动段间距」（`w:beforeAutospacing` / `w:afterAutospacing`）的取值。
/// Word 勾上「自动」后段前/段后显示为 Auto，实际约 14pt（≈0.49cm）；想调只改这一处。
const AUTO_SPACE_PT: f32 = 14.0;

/// 单倍行距折算出行高的经验系数（`w:beforeLines` / `w:afterLines` 的「1 行」基准）。
/// 字号 × 1.2 是 Word 单倍行距在常见中文字体下的典型值。
const SINGLE_LINE_FACTOR: f32 = 1.2;

/// 文档里完全查不到字号时的兜底（中文文档的默认五号字 = 10.5pt）。
/// 只有「字符单位缩进」需要它：1 字符 = 1 个字号。
const DEFAULT_FONT_SIZE_PT: f32 = 10.5;

impl StyleSheet {
    /// `theme` 必须传进来：样式里的 `w:rFonts` 经常**只写主题字体名**（`w:asciiTheme="minorHAnsi"`），
    /// 拿不到主题就等于整篇没有字体名。
    fn parse(xml: &str, theme: &Theme) -> StyleSheet {
        let mut sheet = StyleSheet::default();
        let Some(root) = parse_xml(xml) else {
            return sheet;
        };

        if let Some(defaults) = root.child("docDefaults").or_else(|| root.find_descendant("docDefaults")) {
            if let Some(rpr) = defaults.child("rPrDefault").and_then(|node| node.child("rPr")) {
                sheet.doc_rpr = parse_run_props(rpr, theme);
            }
            if let Some(ppr) = defaults.child("pPrDefault").and_then(|node| node.child("pPr")) {
                sheet.doc_ppr = parse_para_props(ppr);
            }
        }

        for node in root.children_named("style") {
            let Some(id) = node.attr_local("val").or_else(|| node.attr_local("styleId")) else {
                continue;
            };
            let kind = match node.attr_local("type") {
                Some("paragraph") => StyleKind::Paragraph,
                Some("character") => StyleKind::Character,
                _ => StyleKind::Other,
            };
            // 表格样式 / 编号样式不参与文字排版，直接跳过（省内存也避免误用）
            if kind == StyleKind::Other {
                continue;
            }
            if matches!(node.attr_local("default"), Some("1") | Some("true") | Some("on")) {
                match kind {
                    StyleKind::Paragraph => sheet.default_para_style = Some(id.to_string()),
                    StyleKind::Character => sheet.default_char_style = Some(id.to_string()),
                    StyleKind::Other => {}
                }
            }
            let def = StyleDef {
                name: node.child("name").and_then(|node| node.attr_local("val")).map(str::to_string),
                based_on: node.child("basedOn").and_then(|node| node.attr_local("val")).map(str::to_string),
                // 注意：样式里的 rPr 也可能只写主题字体，所以这里必须传主题
                rpr: node
                    .child("rPr")
                    .map(|node| parse_run_props(node, theme))
                    .unwrap_or_default(),
                ppr: node.child("pPr").map(parse_para_props).unwrap_or_default(),
            };
            sheet.styles.insert(id.to_string(), def);
        }
        sheet
    }

    /// 段落样式的继承链：`[本样式, 父样式, 祖父, …]`，最多 [`MAX_STYLE_DEPTH`] 层，防环。
    fn chain(&self, style_id: Option<&str>) -> Vec<&StyleDef> {
        let start = style_id
            .map(str::to_string)
            .or_else(|| self.default_para_style.clone());
        let mut out = Vec::new();
        let mut visited: Vec<String> = Vec::new();
        let mut current = start;
        while let Some(id) = current {
            if out.len() >= MAX_STYLE_DEPTH || visited.iter().any(|seen| seen == &id) {
                break;
            }
            let Some(def) = self.styles.get(&id) else { break };
            visited.push(id);
            out.push(def);
            current = def.based_on.clone();
        }
        out
    }

    /// 段落样式链给出的字符属性（不含 docDefaults，那层另外兜底）
    fn chain_rpr(&self, style_id: Option<&str>) -> RunProps {
        let mut props = RunProps::default();
        for def in self.chain(style_id) {
            props.apply_over(&def.rpr);
        }
        props
    }

    /// 段落样式链给出的段落属性
    fn chain_ppr(&self, style_id: Option<&str>) -> ParaProps {
        let mut props = ParaProps::default();
        for def in self.chain(style_id) {
            props.apply_over(&def.ppr);
        }
        props
    }

    fn style_name(&self, style_id: &str) -> Option<String> {
        self.styles.get(style_id).and_then(|def| def.name.clone())
    }
}

/// 主题（`word/theme/theme1.xml`）里的字体方案。
/// 真实样本的 `docDefaults` 只写了 `w:asciiTheme="minorHAnsi"`，没有具体字体名 ——
/// 不解析主题就整篇拿不到字体，这是"看起来不像"最隐蔽的一种原因。
#[derive(Debug, Clone, Default)]
struct Theme {
    major_latin: Option<String>,
    major_ea: Option<String>,
    minor_latin: Option<String>,
    minor_ea: Option<String>,
}

impl Theme {
    fn parse(xml: &str) -> Theme {
        let Some(root) = parse_xml(xml) else {
            return Theme::default();
        };
        let Some(scheme) = root.find_descendant("fontScheme") else {
            return Theme::default();
        };
        Theme {
            major_latin: scheme.child("majorFont").and_then(|node| typeface(node, "latin")),
            major_ea: scheme.child("majorFont").and_then(east_asian_typeface),
            minor_latin: scheme.child("minorFont").and_then(|node| typeface(node, "latin")),
            minor_ea: scheme.child("minorFont").and_then(east_asian_typeface),
        }
    }

    /// `w:asciiTheme="minorHAnsi"` 这类主题字体名 → 真实字体名
    fn font_for(&self, theme_name: &str) -> Option<String> {
        match theme_name {
            "minorHAnsi" | "minorAscii" | "minorBidi" => self.minor_latin.clone(),
            "majorHAnsi" | "majorAscii" | "majorBidi" => self.major_latin.clone(),
            "minorEastAsia" => self.minor_ea.clone(),
            "majorEastAsia" => self.major_ea.clone(),
            _ => None,
        }
    }
}

fn typeface(font_node: &XmlNode, which: &str) -> Option<String> {
    let value = font_node.child(which)?.attr("typeface")?;
    if value.trim().is_empty() {
        None
    } else {
        Some(value.to_string())
    }
}

/// 东亚字体：`a:ea` 在 Office 主题里**通常是空串**，此时按简体中文用 `<a:font script="Hans">` 兜底。
fn east_asian_typeface(font_node: &XmlNode) -> Option<String> {
    if let Some(name) = typeface(font_node, "ea") {
        return Some(name);
    }
    for script in ["Hans", "Hant", "Jpan", "Hang"] {
        if let Some(node) = font_node
            .children_named("font")
            .find(|node| node.attr("script") == Some(script))
        {
            if let Some(name) = node.attr("typeface") {
                if !name.trim().is_empty() {
                    return Some(name.to_string());
                }
            }
        }
    }
    None
}

/// 用最终生效的字符属性造一个 [`Run`]。
/// 空结果的域（`PAGE`/`NUMPAGES` 没有缓存文本）也要产出 run，所以构造逻辑抽出来共用。
fn run_of(rpr: &RunProps, text: String, style_id: Option<String>, field: Option<FieldKind>) -> Run {
    Run {
        text,
        bold: rpr.bold.unwrap_or(false),
        italic: rpr.italic.unwrap_or(false),
        underline: rpr.underline.unwrap_or(false),
        strike: rpr.strike.unwrap_or(false),
        font: rpr.font_latin.clone(),
        font_east_asia: rpr.font_ea.clone(),
        size_pt: rpr.size_half_pt.map(|half| half / 2.0),
        color: rpr.color.clone(),
        highlight: rpr.highlight.clone(),
        vert_align: rpr.vert_align.clone(),
        style_id,
        field,
    }
}

/// `w:rPr` → 字符属性
fn parse_run_props(node: &XmlNode, theme: &Theme) -> RunProps {
    let mut props = RunProps::default();

    if let Some(fonts) = node.child("rFonts") {
        // 显式字体名优先；只写了主题字体名时查主题表
        props.font_latin = fonts
            .attr_local("ascii")
            .or_else(|| fonts.attr_local("hAnsi"))
            .filter(|value| !value.trim().is_empty())
            .map(str::to_string)
            .or_else(|| {
                fonts
                    .attr_local("asciiTheme")
                    .or_else(|| fonts.attr_local("hAnsiTheme"))
                    .and_then(|name| theme.font_for(name))
            });
        props.font_ea = fonts
            .attr_local("eastAsia")
            .filter(|value| !value.trim().is_empty())
            .map(str::to_string)
            .or_else(|| {
                fonts
                    .attr_local("eastAsiaTheme")
                    .and_then(|name| theme.font_for(name))
            });
    }

    props.bold = on_off(node.child("b"));
    props.italic = on_off(node.child("i"));
    props.strike = on_off(node.child("strike")).or_else(|| on_off(node.child("dstrike")));
    // 下划线：`w:u` 不带 val 就是单下划线；val="none" 才是明确取消
    props.underline = node.child("u").map(|under| {
        !matches!(under.attr_local("val"), Some("none") | Some("0") | Some("false"))
    });

    if let Some(size) = node.child("sz").and_then(|size| number(size.attr_local("val"))) {
        // 半磅 → pt
        if size > 0.0 {
            props.size_half_pt = Some(size);
        }
    }

    if let Some(color) = node.child("color").and_then(|color| color.attr_local("val")) {
        props.color = normalize_color(color);
    }

    // 高亮（文字荧光笔）优先；没有就看底纹填充
    if let Some(highlight) = node.child("highlight").and_then(|node| node.attr_local("val")) {
        props.highlight = highlight_color(highlight);
    }
    if props.highlight.is_none() {
        if let Some(fill) = node.child("shd").and_then(|node| node.attr_local("fill")) {
            props.highlight = normalize_color(fill);
        }
    }

    if let Some(align) = node.child("vertAlign").and_then(|node| node.attr_local("val")) {
        match align {
            "superscript" | "subscript" => props.vert_align = Some(align.to_string()),
            _ => {}
        }
    }

    props
}

/// `w:pPr` → 段落属性
fn parse_para_props(node: &XmlNode) -> ParaProps {
    let mut props = ParaProps::default();

    if let Some(align) = node.child("jc").and_then(|node| node.attr_local("val")) {
        props.align = Some(align.to_string());
    }

    if let Some(indent) = node.child("ind") {
        // 严格 ISO 文档用 start/end，Word 用 left/right
        props.indent_left_pt = twips_to_pt(indent.attr_local("left").or_else(|| indent.attr_local("start")));
        props.indent_right_pt = twips_to_pt(indent.attr_local("right").or_else(|| indent.attr_local("end")));
        // 首行缩进与悬挂缩进互斥：悬挂为负的首行缩进
        let hanging_twips = twips_to_pt(indent.attr_local("hanging"));
        if let Some(hanging) = hanging_twips {
            props.first_line_pt = Some(-hanging);
        } else {
            props.first_line_pt = twips_to_pt(indent.attr_local("firstLine"));
        }
        // 字符单位（百分之一字符）：层叠结束后按最终字号折算，见 ParaProps::resolve_char_indents
        props.left_chars = number(indent.attr_local("leftChars").or_else(|| indent.attr_local("startChars")));
        props.right_chars = number(indent.attr_local("rightChars").or_else(|| indent.attr_local("endChars")));
        props.first_line_chars = match number(indent.attr_local("hangingChars")) {
            Some(hanging) => Some(-hanging),
            None => number(indent.attr_local("firstLineChars")),
        };
    }

    if let Some(spacing) = node.child("spacing") {
        props.space_before_pt = twips_to_pt(spacing.attr_local("before"));
        props.space_after_pt = twips_to_pt(spacing.attr_local("after"));
        // 行单位（百分之一行）：层叠结束后按最终行高折算（缺 twips 时才生效）
        props.before_lines = number(spacing.attr_local("beforeLines"));
        props.after_lines = number(spacing.attr_local("afterLines"));
        // 自动段间距：`w:beforeAutospacing` / `w:afterAutospacing` 是 w:spacing 的**属性**；
        // 先记开关，层叠结束后定稿（见 ParaProps::resolve_spacing）
        props.auto_before = attr_on_off(spacing, "beforeAutospacing");
        props.auto_after = attr_on_off(spacing, "afterAutospacing");
        if let Some(line) = number(spacing.attr_local("line")) {
            let rule = spacing.attr_local("lineRule").unwrap_or("auto");
            let value = match rule {
                // auto：240 = 单倍行距
                "auto" => line / 240.0,
                _ => line / 20.0, // exact / atLeast：twip → pt
            };
            props.line = Some(LineSpacing {
                kind: match rule {
                    "auto" => "multiple".to_string(),
                    other => other.to_string(),
                },
                value,
            });
        }
    }

    if let Some(level) = node
        .child("outlineLvl")
        .and_then(|node| number(node.attr_local("val")))
    {
        if (0.0..=8.0).contains(&level) {
            props.outline_level = Some(level as u8);
        }
    }

    // 段落边框（合同/表单里"空段落 + 下边框"就是一条横线）
    props.borders = node.child("pBdr").and_then(parse_paragraph_borders);

    props
}

/// `w:pBdr` → 段落边框。`w:val` 为 `none`/`nil` 的边不产出（是"明确不要边框"）；
/// 四边都没有就返回 `None`（前端少一层判断）。
fn parse_paragraph_borders(node: &XmlNode) -> Option<ParagraphBorders> {
    let side = |name: &str| -> Option<BorderSpec> {
        let border = node.child(name)?;
        let style = border.attr_local("val").unwrap_or("single").trim();
        if style.is_empty() || style.eq_ignore_ascii_case("none") || style.eq_ignore_ascii_case("nil") {
            return None;
        }
        Some(BorderSpec {
            style: style.to_string(),
            // `w:sz` 是八分之一磅
            width_pt: number_f64(border.attr_local("sz")).unwrap_or(0.0) / 8.0,
            // 边框颜色只把 "auto" 当没写 —— 白色是**合法的可见设置**（等于隐形线），
            // 不能像底纹那样把 FFFFFF 也吃掉
            color: border.attr_local("color").and_then(normalize_color_strict),
            // `w:space` 的单位就是 pt
            space_pt: number_f64(border.attr_local("space")),
        })
    };
    let borders = ParagraphBorders {
        top: side("top"),
        bottom: side("bottom"),
        left: side("left"),
        right: side("right"),
    };
    if borders.top.is_none() && borders.bottom.is_none() && borders.left.is_none() && borders.right.is_none()
    {
        return None;
    }
    Some(borders)
}

/// `<w:b/>` / `<w:b w:val="0"/>` → Some(true/false)；节点不存在 → None（没表态）
fn on_off(node: Option<&XmlNode>) -> Option<bool> {
    let node = node?;
    match node.attr_local("val") {
        None => Some(true),
        Some("0") | Some("false") | Some("off") => Some(false),
        Some(_) => Some(true),
    }
}

/// **属性**形式的开关（`w:beforeAutospacing="1"`）；属性不存在 → None（没表态）
fn attr_on_off(node: &XmlNode, name: &str) -> Option<bool> {
    match node.attr_local(name) {
        None => None,
        Some("0") | Some("false") | Some("off") => Some(false),
        Some(_) => Some(true),
    }
}

fn number(value: Option<&str>) -> Option<f32> {
    value?.trim().parse::<f32>().ok()
}

/// f64 版（页面几何用它，避免 pt 值来回折损精度）
fn number_f64(value: Option<&str>) -> Option<f64> {
    value?.trim().parse::<f64>().ok()
}

fn integer(value: Option<&str>) -> Option<i64> {
    value?.trim().parse::<i64>().ok()
}

/// twip（1/20 pt）→ pt
fn twips_to_pt(value: Option<&str>) -> Option<f32> {
    number(value).map(|twips| twips / 20.0)
}

/// twip → pt（f64 版，页面几何用）
fn twips_to_pt_f64(twips: f64) -> f64 {
    twips / 20.0
}

/// twip → px（96 dpi：1pt = 4/3 px）
fn twips_to_px(value: f32) -> f32 {
    value / 20.0 * 4.0 / 3.0
}

/// 1 英寸 = 914400 EMU = 96 px → 9525 EMU/px
fn emu_to_px(value: f32) -> f32 {
    value / 9525.0
}

/// 1 pt = 12700 EMU（形状几何用它）
const EMU_PER_PT: f64 = 12700.0;

/// `w:color` 值 → RRGGBB；`auto` / 非法值当没写
fn normalize_color(value: &str) -> Option<String> {
    let value = value.trim();
    if value.is_empty() || value.eq_ignore_ascii_case("auto") {
        return None;
    }
    // `w:shd w:val="clear" w:fill="FFFFFF"` 是"没有底纹"的惯用写法，别把整篇刷成白色
    if value.eq_ignore_ascii_case("FFFFFF") {
        return None;
    }
    let hex: String = value
        .trim_start_matches('#')
        .chars()
        .filter(char::is_ascii_hexdigit)
        .collect();
    match hex.len() {
        6 => Some(hex.to_uppercase()),
        3 => Some(hex.chars().flat_map(|c| [c, c]).collect::<String>().to_uppercase()),
        _ => None,
    }
}

/// `w:highlight` 的具名颜色 → RRGGBB
fn highlight_color(name: &str) -> Option<String> {
    let hex = match name {
        "black" => "000000",
        "blue" => "0000FF",
        "cyan" => "00FFFF",
        "darkBlue" => "000080",
        "darkCyan" => "008080",
        "darkGray" => "808080",
        "darkGreen" => "008000",
        "darkMagenta" => "800080",
        "darkRed" => "800000",
        "darkYellow" => "808000",
        "green" => "00FF00",
        "lightGray" => "C0C0C0",
        "magenta" => "FF00FF",
        "red" => "FF0000",
        "white" => "FFFFFF",
        "yellow" => "FFFF00",
        "none" => return None,
        other => return normalize_color(other),
    };
    Some(hex.to_string())
}

/// **严格**颜色解析：只认 `#RRGGBB` / `RRGGBB` / `RGB` 与 VML 的常用具名色，
/// 只把 `auto` 当"没写"。
///
/// 与 [`normalize_color`] 的区别：那个会把 `FFFFFF` 也当成"没写"（底纹的惯用写法），
/// 还会把非十六进制字符**过滤掉**（`"black"` 会被削成 `bac` → 变成 `BBAACC`！）。
/// 边框色与 VML 的 `strokecolor` 都用这个，避免把白线变没、把颜色名解析成怪色。
fn normalize_color_strict(value: &str) -> Option<String> {
    let value = value.trim();
    if value.is_empty() || value.eq_ignore_ascii_case("auto") || value.eq_ignore_ascii_case("none") {
        return None;
    }
    let lowered = value.to_ascii_lowercase();
    if let Some(hex) = match lowered.as_str() {
        "black" | "windowtext" => Some("000000"),
        "white" | "window" => Some("FFFFFF"),
        "red" => Some("FF0000"),
        "green" => Some("008000"),
        "lime" => Some("00FF00"),
        "blue" => Some("0000FF"),
        "yellow" => Some("FFFF00"),
        "gray" | "grey" => Some("808080"),
        "silver" => Some("C0C0C0"),
        "maroon" => Some("800000"),
        "navy" => Some("000080"),
        "teal" => Some("008080"),
        "purple" => Some("800080"),
        "olive" => Some("808000"),
        "fuchsia" | "magenta" => Some("FF00FF"),
        "aqua" | "cyan" => Some("00FFFF"),
        _ => None,
    } {
        return Some(hex.to_string());
    }
    let hex = value.trim_start_matches('#');
    let is_hex = hex.chars().all(|ch| ch.is_ascii_hexdigit()) && !hex.is_empty();
    match (is_hex, hex.len()) {
        (true, 6) => Some(hex.to_uppercase()),
        (true, 3) => Some(hex.chars().flat_map(|ch| [ch, ch]).collect::<String>().to_uppercase()),
        _ => None,
    }
}

/* ================================================================================== */
/* 四、编号与多级列表（numbering.xml）                                                */
/* ================================================================================== */

/// 某一级编号的定义
#[derive(Debug, Clone)]
struct LevelDef {
    /// `w:start`（默认 1）
    start: u32,
    /// `w:numFmt`（decimal / bullet / chineseCounting …）
    format: String,
    /// `w:lvlText`，如 `%1.` / `（%1）` / `•`
    text: String,
    /// `w:suff`：标记与正文的间隔（tab / space / nothing）
    suffix: String,
    /// 该级自带的缩进（作为段落属性的中间层）
    indent_left_twips: Option<f32>,
    /// 悬挂缩进（正数，渲染时取负作为首行缩进）
    hanging_twips: Option<f32>,
    /// `w:leftChars`（百分之一字符）：层叠后按最终字号折算
    left_chars: Option<f32>,
    /// `w:firstLineChars`（正）或 `-w:hangingChars`（负）：同样按最终字号折算
    first_line_chars: Option<f32>,
}

impl Default for LevelDef {
    fn default() -> Self {
        // 文档里找不到级别定义时的兜底：按最普通的 "1." 处理，绝不让编号整段消失
        LevelDef {
            start: 1,
            format: "decimal".to_string(),
            text: "%1.".to_string(),
            suffix: "tab".to_string(),
            indent_left_twips: None,
            hanging_twips: None,
            left_chars: None,
            first_line_chars: None,
        }
    }
}

#[derive(Debug, Default)]
struct AbstractNum {
    /// ilvl → 级别定义（0..8）
    levels: HashMap<u8, LevelDef>,
}

#[derive(Debug, Default)]
struct NumDef {
    abstract_id: u32,
    /// `w:lvlOverride/w:lvl`：整级覆盖（WPS/Word 重排编号时会出现）
    level_overrides: HashMap<u8, LevelDef>,
    /// `w:lvlOverride/w:startOverride`：只覆盖起始值
    start_overrides: HashMap<u8, u32>,
}

#[derive(Debug, Default)]
struct Numbering {
    nums: HashMap<u32, NumDef>,
    abstracts: HashMap<u32, AbstractNum>,
}

impl Numbering {
    fn parse(xml: &str) -> Numbering {
        let mut numbering = Numbering::default();
        let Some(root) = parse_xml(xml) else {
            return numbering;
        };

        for node in root.children_named("abstractNum") {
            let Some(id) = integer(node.attr_local("abstractNumId")) else {
                continue;
            };
            let mut abstract_num = AbstractNum::default();
            for level in node.children_named("lvl") {
                let index = integer(level.attr_local("ilvl")).unwrap_or(0).clamp(0, 8) as u8;
                abstract_num.levels.insert(index, parse_level(level));
            }
            numbering.abstracts.insert(id.max(0) as u32, abstract_num);
        }

        for node in root.children_named("num") {
            let Some(id) = integer(node.attr_local("numId")) else {
                continue;
            };
            let mut def = NumDef {
                abstract_id: node
                    .child("abstractNumId")
                    .and_then(|node| integer(node.attr_local("val")))
                    .unwrap_or(0)
                    .max(0) as u32,
                ..NumDef::default()
            };
            for over in node.children_named("lvlOverride") {
                let Some(index) = integer(over.attr_local("ilvl")) else {
                    continue;
                };
                let index = index.clamp(0, 8) as u8;
                if let Some(level) = over.child("lvl") {
                    def.level_overrides.insert(index, parse_level(level));
                }
                if let Some(start) = over
                    .child("startOverride")
                    .and_then(|node| integer(node.attr_local("val")))
                {
                    def.start_overrides.insert(index, start.max(0) as u32);
                }
            }
            numbering.nums.insert(id.max(0) as u32, def);
        }
        numbering
    }

    /// 取某一级的定义（级别覆盖 > 抽象编号；都没有就兜底）
    fn level(&self, num_id: u32, ilvl: u8) -> (LevelDef, u32) {
        let Some(def) = self.nums.get(&num_id) else {
            return (LevelDef::default(), 1);
        };
        let mut level = def
            .level_overrides
            .get(&ilvl)
            .cloned()
            .or_else(|| {
                self.abstracts
                    .get(&def.abstract_id)
                    .and_then(|abs| abs.levels.get(&ilvl))
                    .cloned()
            })
            .unwrap_or_default();
        if let Some(start) = def.start_overrides.get(&ilvl) {
            level.start = *start;
        }
        (level, def.abstract_id)
    }

    /// 该级自带的缩进 → 段落属性（层叠里位于"段落样式之上、直接格式之下"）
    fn level_indent(&self, num_id: u32, ilvl: u8) -> ParaProps {
        let (level, _) = self.level(num_id, ilvl);
        let mut props = ParaProps::default();
        props.indent_left_pt = level.indent_left_twips.map(|twips| twips / 20.0);
        props.first_line_pt = level.hanging_twips.map(|twips| -twips / 20.0);
        // 有些生成器在编号级别上只写 Chars：一起带上，等层叠结束按最终字号折算
        props.left_chars = level.left_chars;
        props.first_line_chars = level.first_line_chars;
        props
    }
}

fn parse_level(node: &XmlNode) -> LevelDef {
    let mut level = LevelDef::default();
    if let Some(start) = integer(node.child("start").and_then(|node| node.attr_local("val"))) {
        level.start = start.max(0) as u32;
    }
    if let Some(format) = node.child("numFmt").and_then(|node| node.attr_local("val")) {
        level.format = format.to_string();
    }
    if let Some(text) = node.child("lvlText").and_then(|node| node.attr_local("val")) {
        level.text = text.to_string();
    }
    if let Some(suffix) = node.child("suff").and_then(|node| node.attr_local("val")) {
        level.suffix = suffix.to_string();
    }
    if let Some(indent) = node.child("pPr").and_then(|node| node.child("ind")) {
        level.indent_left_twips = number(indent.attr_local("left").or_else(|| indent.attr_local("start")));
        level.hanging_twips = number(
            indent
                .attr_local("hanging")
                .or_else(|| indent.attr_local("firstLine")),
        );
        // 字符单位同理（编号级别的缩进也可能是「2 字符」）
        level.left_chars = number(
            indent
                .attr_local("leftChars")
                .or_else(|| indent.attr_local("startChars")),
        );
        level.first_line_chars = match number(indent.attr_local("hangingChars")) {
            Some(hanging) => Some(-hanging),
            None => number(indent.attr_local("firstLineChars")),
        };
    }
    level
}

/// 编号计数器：`(numId, ilvl) → 当前值`。
///
/// 为什么要按 numId 分别计数：`w:num` 才是"一套独立的编号"（同一 abstractNum 可以被多个
/// numId 复用，各自从头数），所以计数器必须挂在 numId 上而不是 abstractNumId 上，
/// 否则两个都指向 abstractNum 8 的列表会被串成一条连续编号。
#[derive(Debug, Default)]
struct NumberingState {
    counters: HashMap<(u32, u8), u32>,
}

impl NumberingState {
    /// 推进一级并算出该段要显示的前缀文本
    fn next(&mut self, numbering: &Numbering, num_id: u32, ilvl: u8) -> ListInfo {
        let (level, _) = numbering.level(num_id, ilvl);
        let start = level.start.max(1);

        let counter = self.counters.entry((num_id, ilvl)).or_insert(start - 1);
        *counter = counter.saturating_add(1);

        // 更深的级别重新计数：1.1 之后的 1.2 出现时，2.1 必须从头开始
        self.counters
            .retain(|(other_num, other_level), _| !(*other_num == num_id && *other_level > ilvl));

        let prefix = self.render(&level.text, numbering, num_id);
        let ordered = !matches!(level.format.as_str(), "bullet" | "none");
        ListInfo {
            num_id,
            level: ilvl,
            ordered,
            prefix,
            format: level.format.clone(),
            suffix: level.suffix.clone(),
        }
    }

    /// `%N` → 第 N 级（1 起）的编号文本；父级没用过时按它的 start 显示。
    /// 项目符号级的 `lvlText` 就是符号本身（`•` / `` / `o`），不含 `%N`，会原样带出。
    fn render(&self, text: &str, numbering: &Numbering, num_id: u32) -> String {
        let mut out = String::with_capacity(text.len() + 4);
        let chars: Vec<char> = text.chars().collect();
        let mut index = 0;
        while index < chars.len() {
            let ch = chars[index];
            if ch == '%' && index + 1 < chars.len() && chars[index + 1].is_ascii_digit() {
                let level_index = chars[index + 1] as u8 - b'0';
                if (1..=9).contains(&level_index) {
                    let target = level_index - 1;
                    let (level, _) = numbering.level(num_id, target);
                    let value = self
                        .counters
                        .get(&(num_id, target))
                        .copied()
                        // 父级还没出现过：按它的起始值显示（与 Word 的观感一致）
                        .unwrap_or_else(|| level.start.max(1));
                    out.push_str(&format_value(&level.format, value));
                    index += 2;
                    continue;
                }
            }
            out.push(ch);
            index += 1;
        }
        out
    }
}

/// 编号格式 → 文本。未知格式退回阿拉伯数字（宁可显示成 "1." 也不要空着）。
fn format_value(format: &str, value: u32) -> String {
    match format {
        "bullet" | "none" => String::new(),
        "decimal" => value.to_string(),
        "decimalZero" => format!("{value:02}"),
        "lowerLetter" => latin_letters(value, false),
        "upperLetter" => latin_letters(value, true),
        "lowerRoman" => roman(value, false),
        "upperRoman" => roman(value, true),
        "ordinal" => ordinal(value),
        "decimalFullWidth" => value
            .to_string()
            .chars()
            .map(|ch| char::from_u32(0xFF10 + (ch as u32 - '0' as u32)).unwrap_or(ch))
            .collect(),
        "decimalEnclosedCircle" | "decimalEnclosedCircleChinese" => enclosed(value, 0x2460),
        "decimalEnclosedParen" => enclosed(value, 0x2474),
        "decimalEnclosedFullstop" => enclosed(value, 0x2488),
        "chineseCounting"
        | "chineseCountingThousand"
        | "chineseLegalSimplified"
        | "ideographDigital"
        | "ideographTraditional"
        | "japaneseCounting"
        | "japaneseDigitalTenThousand"
        | "japaneseLegal"
        | "koreanCounting"
        | "koreanDigital"
        | "taiwaneseCounting"
        | "taiwaneseCountingThousand" => cjk_number(value),
        _ => value.to_string(),
    }
}

/// 1 → a、26 → z、27 → aa（Word 的 lowerLetter/upperLetter 规则）
fn latin_letters(value: u32, upper: bool) -> String {
    let mut value = value.max(1);
    let mut out = Vec::new();
    while value > 0 {
        let rem = ((value - 1) % 26) as u8;
        out.push((b'a' + rem) as char);
        value = (value - 1) / 26;
    }
    out.reverse();
    let text: String = out.into_iter().collect();
    if upper {
        text.to_uppercase()
    } else {
        text
    }
}

fn roman(value: u32, upper: bool) -> String {
    const TABLE: [(u32, &str); 13] = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut value = value.min(3999);
    let mut out = String::new();
    for (amount, symbol) in TABLE {
        while value >= amount {
            out.push_str(symbol);
            value -= amount;
        }
    }
    if upper {
        out.to_uppercase()
    } else {
        out
    }
}

fn ordinal(value: u32) -> String {
    let suffix = if (11..=13).contains(&(value % 100)) {
        "th"
    } else {
        match value % 10 {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th",
        }
    };
    format!("{value}{suffix}")
}

/// 带圈/带括号数字（① ⑴ ⒈）：Unicode 块里只排到 20，超出就退回普通数字
fn enclosed(value: u32, base: u32) -> String {
    if (1..=20).contains(&value) {
        if let Some(ch) = char::from_u32(base + value - 1) {
            return ch.to_string();
        }
    }
    value.to_string()
}

/// 中文数字（一、十、二十三、一百零五…）；用于 chineseCounting / japaneseCounting 等。
/// 这些格式在 Word 里的差异只在"万以上怎么读"，本项目只到千位，取最通用的一套。
fn cjk_number(value: u32) -> String {
    const DIGITS: [char; 10] = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
    const UNITS: [char; 4] = [' ', '十', '百', '千'];
    if value == 0 {
        return "零".to_string();
    }
    if value >= 10000 {
        // 万以上：递归处理"万"位（编号极少超过 9999，这里只是别显示成乱码）
        let high = value / 10000;
        let low = value % 10000;
        let mut out = cjk_number(high);
        out.push('万');
        if low > 0 {
            if low < 1000 {
                out.push('零');
            }
            out.push_str(&cjk_number(low));
        }
        return out;
    }
    let digits: Vec<u32> = value
        .to_string()
        .chars()
        .filter_map(|ch| ch.to_digit(10))
        .collect();
    let mut out = String::new();
    let len = digits.len();
    for (index, digit) in digits.iter().enumerate() {
        let unit_index = len - index - 1;
        if *digit == 0 {
            // 连续的 0 只写一个"零"，且末尾的 0 不写
            if !out.ends_with('零') && index + 1 < len {
                out.push('零');
            }
            continue;
        }
        if !(*digit == 1 && unit_index == 1 && index == 0) {
            // 十、十一 里的"一十"按中文习惯简写成"十"
            out.push(DIGITS[*digit as usize]);
        }
        if unit_index > 0 {
            out.push(UNITS[unit_index]);
        }
    }
    if out.is_empty() {
        "零".to_string()
    } else {
        out
    }
}

/* ================================================================================== */
/* 五、关系与媒体（document.xml.rels）                                                */
/* ================================================================================== */

#[derive(Debug, Clone, Default)]
struct RelInfo {
    /// 解析后的包内路径（`media/image1.png` → `word/media/image1.png`）
    target: String,
    /// 关系类型（URI 末段：image / chart / oleObject / hyperlink …）
    kind: String,
    /// `TargetMode="External"`：目标是外部 URL，不在包里
    external: bool,
}

/// 解析 `word/_rels/document.xml.rels`
fn parse_rels(xml: &str) -> HashMap<String, RelInfo> {
    let mut out = HashMap::new();
    let Some(root) = parse_xml(xml) else {
        return out;
    };
    for node in root.children_named("Relationship") {
        let Some(id) = node.attr("Id") else { continue };
        let target = node.attr("Target").unwrap_or_default();
        let kind = node
            .attr("Type")
            .map(|uri| uri.rsplit('/').next().unwrap_or(uri).to_string())
            .unwrap_or_default();
        let external = node.attr("TargetMode") == Some("External");
        out.insert(
            id.to_string(),
            RelInfo {
                target: if external {
                    target.to_string()
                } else {
                    resolve_part_path("word", target)
                },
                kind,
                external,
            },
        );
    }
    out
}

/// `Target` 相对部件目录的路径 → 包内绝对路径（`../customXml/item1.xml` → `customXml/item1.xml`）
fn resolve_part_path(base_dir: &str, target: &str) -> String {
    let target = target.replace('\\', "/");
    let joined = if let Some(absolute) = target.strip_prefix('/') {
        absolute.to_string()
    } else if base_dir.is_empty() {
        target.clone()
    } else {
        format!("{base_dir}/{target}")
    };
    let mut parts: Vec<&str> = Vec::new();
    for segment in joined.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/* ================================================================================== */
/* 六、块构建器                                                                       */
/* ================================================================================== */

/// 单元格没有明确尺寸时的图片兜底显示尺寸（px）—— 宁可给个可见的占位尺寸，
/// 也不要算出 0×0 让图片"消失"。
const FALLBACK_IMAGE_PX: (f32, f32) = (320.0, 200.0);

/// 段落里的行内内容收集结果
#[derive(Default)]
struct ParagraphContent {
    runs: Vec<Run>,
    /// 段落里的**非文字块**（图片 / 形状 / 占位对象），按文档顺序排好。
    /// 合成一个列表而不是分几个数组：这样"图片 - 线 - 占位"的先后顺序不会被打乱。
    objects: Vec<Block>,
    /// 段落里出现过 `<w:br w:type="page"/>`
    page_break: bool,
    /// 域（`PAGE` / `NUMPAGES`）的状态：Word 把 begin / instrText / separate / end
    /// 拆在**多个 run** 里，所以要跨 run 记状态
    field: FieldState,
}

/// 域解析状态。用**栈**做嵌套域的深度计数（`IF` 里套 `PAGE` 很常见）。
#[derive(Default)]
struct FieldState {
    /// 每个未结束的域一层：是否已经 `separate`（之后的文本就是缓存结果）+ 指令原文
    frames: Vec<FieldFrame>,
    /// `w:fldSimple`（属性式域，`w:instr=" PAGE "`）：里面的 run 全是缓存结果
    simple: Option<FieldKind>,
}

#[derive(Default)]
struct FieldFrame {
    /// 已经遇到 `separate`：后面的文本就是缓存结果
    separated: bool,
    /// 指令文本（`w:instrText` 可能跨若干 run）
    instruction: String,
    /// 结果区里出现过非空的 `w:t` 吗（没有就要补一个带标记的空 run，见 `on_fld_char`）
    text_seen: bool,
}

impl FieldState {
    /// `w:fldChar@w:fldCharType`：begin 入栈 / separate 标记 / end 出栈
    /// （注意属性名是 `w:fldCharType`，本地名不是 `type`）。
    ///
    /// `end` 时如果这个域**已识别（PAGE/NUMPAGES）、已经算过（separate）、
    /// 但结果区里一个字的缓存结果都没有**，返回它的类型：调用方要补一个带标记的空 run。
    /// 否则渲染器没有可替换的对象 —— 新建、或从未打印过的文档，页脚页码会是一片空白。
    fn on_fld_char(&mut self, node: &XmlNode) -> Option<FieldKind> {
        match node.attr_local("fldCharType").map(str::trim) {
            Some("begin") => self.frames.push(FieldFrame::default()),
            Some("separate") => {
                if let Some(frame) = self.frames.last_mut() {
                    frame.separated = true;
                }
            }
            Some("end") => {
                // 多出来的 end（XML 有毛病）不该把已有状态清空
                let frame = self.frames.pop()?;
                if frame.separated && !frame.text_seen {
                    return field_kind_from_instruction(&frame.instruction);
                }
            }
            _ => {}
        }
        None
    }

    /// `w:instrText`：拼到最近一层未结束的域的指令里
    fn on_instruction(&mut self, text: &str) {
        if let Some(frame) = self.frames.last_mut() {
            frame.instruction.push_str(text);
        }
    }

    /// 结果区里读到了非空的 `w:t`：给所有"已算过"的域记一笔
    /// （嵌套时外层域也不该再补一个空 run）
    fn note_text(&mut self) {
        for frame in self.frames.iter_mut().filter(|frame| frame.separated) {
            frame.text_seen = true;
        }
    }

    /// 当前文本是不是"会变的域"的缓存结果。
    /// 从最内层往外找第一个**已经 separate** 的域：它的指令决定这段文本的性质 ——
    /// 是 `PAGE`/`NUMPAGES` 就标记，是别的域（`DATE`/`REF`…）就返回 `None`（行为完全不变）。
    fn result_kind(&self) -> Option<FieldKind> {
        if self.simple.is_some() {
            return self.simple;
        }
        for frame in self.frames.iter().rev() {
            if frame.separated {
                return field_kind_from_instruction(&frame.instruction);
            }
        }
        None
    }
}

/// 域指令 → 域类型。指令形如 `PAGE`、` PAGE `、`PAGE \* MERGEFORMAT`、`NUMPAGES \* Arabic`：
/// 取**第一个词**判断、大小写不敏感；只认 `PAGE` 与 `NUMPAGES`，其它一律 `None`。
fn field_kind_from_instruction(instruction: &str) -> Option<FieldKind> {
    let word = instruction.split_whitespace().next()?;
    match word.to_ascii_uppercase().as_str() {
        "PAGE" => Some(FieldKind::Page),
        "NUMPAGES" => Some(FieldKind::NumPages),
        _ => None,
    }
}

impl ParagraphContent {
    fn is_empty(&self) -> bool {
        self.runs.is_empty() && self.objects.is_empty() && !self.page_break
    }
}

struct DocBuilder<'a> {
    styles: &'a StyleSheet,
    numbering: &'a Numbering,
    theme: &'a Theme,
    rels: &'a HashMap<String, RelInfo>,
    numbering_state: NumberingState,
    /// 文档里的 `w:p` 元素总数（含表格单元格内的），用于与 `document_info` 对账
    paragraph_count: usize,
    /// 正文栏宽（pt）= 纸张宽 − 左右页边距：给"整栏宽"的对象用（`v:hr` 水平线）
    column_width_pt: f64,
    /// 当前段落的内容宽度（pt）= 栏宽 − 本段左右缩进（每次解析段落时更新）
    current_width_pt: f64,
}

impl<'a> DocBuilder<'a> {
    fn new(
        styles: &'a StyleSheet,
        numbering: &'a Numbering,
        theme: &'a Theme,
        rels: &'a HashMap<String, RelInfo>,
        page: Option<&PageGeometry>,
    ) -> Self {
        // 没有 w:sectPr 时按 A4 + 默认边距估栏宽（前端也是同一套兜底）
        let column_width_pt = page
            .map(|page| page.width_pt - page.margin_left_pt - page.margin_right_pt)
            .unwrap_or(A4_WIDTH_PT - 2.0 * A4_MARGIN_PT)
            .max(1.0);
        DocBuilder {
            styles,
            numbering,
            theme,
            rels,
            numbering_state: NumberingState::default(),
            paragraph_count: 0,
            column_width_pt,
            current_width_pt: column_width_pt,
        }
    }

    /// 解析一段独立"故事"（页眉 / 页脚）：编号计数器**进出都还原**，
    /// 而且故事里从零开始计数 —— 页眉里的"1."就是 1，不继承正文、也不影响正文。
    fn parse_story(&mut self, container: &XmlNode) -> Vec<Block> {
        let saved = std::mem::take(&mut self.numbering_state.counters);
        let mut blocks = Vec::new();
        self.collect_blocks(container, &mut blocks);
        self.numbering_state.counters = saved;
        blocks
    }

    /// 容器（body / tc / sdtContent）的直接子元素 → 有序块
    fn collect_blocks(&mut self, container: &XmlNode, out: &mut Vec<Block>) {
        for child in container.children() {
            match child.local() {
                "p" => self.push_paragraph(child, out),
                "tbl" => out.push(self.build_table(child)),
                "sdt" => {
                    // 内容控件：真正的内容在 w:sdtContent 里，直接展开，别把文字丢了
                    if let Some(content) = child.child("sdtContent") {
                        self.collect_blocks(content, out);
                    }
                }
                "altChunk" => {
                    let detail = child
                        .attr_local("id")
                        .map(|id| format!("引用的外部内容关系：{id}"))
                        .unwrap_or_else(|| "文档引用了未嵌入的外部内容".to_string());
                    out.push(unsupported("嵌入的外部内容（altChunk，暂不支持显示）", detail));
                }
                // 纯元数据/锚点：跳过不算丢失
                "sectPr" | "bookmarkStart" | "bookmarkEnd" | "proofErr" | "permStart" | "permEnd"
                | "commentRangeStart" | "commentRangeEnd" | "customXmlInsRangeStart"
                | "customXmlInsRangeEnd" | "customXmlDelRangeStart" | "customXmlDelRangeEnd"
                | "customXmlMoveFromRangeStart" | "customXmlMoveFromRangeEnd" => {}
                // 修订：显示最终态 —— w:ins 里的内容要显示，w:del 里的不显示
                "ins" | "moveTo" | "smartTag" | "customXml" => self.collect_blocks(child, out),
                "del" | "moveFrom" => {}
                _ => {
                    // 不认识的容器：只要它直接装着段落/表格就展开（WPS 常见各种包装节点）
                    if child
                        .children()
                        .any(|grand| matches!(grand.local(), "p" | "tbl"))
                    {
                        self.collect_blocks(child, out);
                    }
                }
            }
        }
    }

    /// 一个 `w:p` → 1 个或多个块（图片/占位对象会单独成块）
    fn push_paragraph(&mut self, node: &XmlNode, out: &mut Vec<Block>) {
        self.paragraph_count += 1;

        let style_id = node
            .child("pPr")
            .and_then(|ppr| ppr.child("pStyle"))
            .and_then(|style| style.attr_local("val"))
            .map(str::to_string);

        // ---- 编号：先把 numId/ilvl 读出来，编号级别的缩进要参与段落属性层叠 ----
        let num_pr = node.child("pPr").and_then(|ppr| ppr.child("numPr"));
        let num_id = num_pr
            .and_then(|num| num.child("numId"))
            .and_then(|num| integer(num.attr_local("val")));
        let ilvl = num_pr
            .and_then(|num| num.child("ilvl"))
            .and_then(|num| integer(num.attr_local("val")))
            .unwrap_or(0)
            .clamp(0, 8) as u8;

        // numId=0 是"取消编号"（Word 用 0 表示无编号）
        let list = match num_id {
            Some(id) if id > 0 => Some(
                self.numbering_state
                    .next(self.numbering, id.max(0) as u32, ilvl),
            ),
            _ => None,
        };
        let numbering_ppr = match &list {
            Some(info) => self.numbering.level_indent(info.num_id, info.level),
            None => ParaProps::default(),
        };

        // ---- 段落属性层叠：直接格式 → 编号级别缩进 → 段落样式（basedOn 逐级）→ docDefaults ----
        let direct_ppr = node.child("pPr").map(parse_para_props).unwrap_or_default();
        // 段落自己显式写了段前/段后吗？自动段间距遇到显式值要让位（见 resolve_autospacing）
        let explicit_before = direct_ppr.space_before_pt.is_some();
        let explicit_after = direct_ppr.space_after_pt.is_some();
        let mut ppr = direct_ppr;
        ppr.apply_over(&numbering_ppr);
        ppr.apply_over(&self.styles.chain_ppr(style_id.as_deref()));
        ppr.apply_over(&self.styles.doc_ppr);

        // ---- 字符属性基线（给没有自己 rPr 的 run 用）：段落标记的 rPr → 段落样式链 → docDefaults
        //      `w:pPr/w:rPr` 是"段落标记"的格式，Word 把它当作本段文字默认格式；
        //      很多文档（尤其 WPS）靠它给整段设字体字号，漏掉就会"字号不对"。----
        let mut base_rpr = node
            .child("pPr")
            .and_then(|ppr| ppr.child("rPr"))
            .map(|rpr| parse_run_props(rpr, self.theme))
            .unwrap_or_default();
        base_rpr.apply_over(&self.styles.chain_rpr(style_id.as_deref()));
        base_rpr.apply_over(&self.styles.doc_rpr);

        // ---- 两类"必须等层叠结束"的段落设置在这里定稿 ----
        // 1) 字符单位缩进：1 字符 = 本段最终生效的字号（段落标记 rPr → 样式链 → docDefaults）
        let font_size_pt = base_rpr
            .size_half_pt
            .map(|half| half / 2.0)
            .unwrap_or(DEFAULT_FONT_SIZE_PT);
        ppr.resolve_char_indents(font_size_pt);
        // 2) 段前/段后：显式 twips → 行单位折算 → 自动间距（Auto）→ 层叠值
        ppr.resolve_spacing(font_size_pt, explicit_before, explicit_after);
        // 本段可用的内容宽度（pt）：给 `v:hr` 这类"整栏宽"的形状用
        self.current_width_pt = (self.column_width_pt
            - ppr.indent_left_pt.unwrap_or(0.0) as f64
            - ppr.indent_right_pt.unwrap_or(0.0) as f64)
            .max(1.0);

        // ---- 行内内容 ----
        let mut content = ParagraphContent::default();
        for child in node.children() {
            self.walk_inline(child, &base_rpr, None, &mut content);
        }

        let text: String = content.runs.iter().map(|run| run.text.as_str()).collect();
        let has_visible_text = text.chars().any(|ch| !ch.is_whitespace());

        let page_break_before = node
            .child("pPr")
            .and_then(|ppr| ppr.child("pageBreakBefore"))
            .map(|flag| match flag.attr_local("val") {
                Some("0") | Some("false") | Some("off") => false,
                _ => true,
            })
            .unwrap_or(false);
        // 段落级别的分节符（分节符藏在段落属性里，前端据此画"分页提示线"）
        let section_break = node
            .child("pPr")
            .and_then(|ppr| ppr.child("sectPr"))
            .map(|sect| {
                sect.child("type")
                    .and_then(|node| node.attr_local("val"))
                    .unwrap_or("nextPage")
                    .to_string()
            });

        // 只含分页符的段落（Word 里 Ctrl+Enter 就是这个形态）→ 独立的分页块。
        // 注意判断的是"没有可见文字"而不是"没有 run"：run 里的 `\n` 是分页符本身带来的。
        if !has_visible_text && content.objects.is_empty() && content.page_break {
            out.push(Block::PageBreak);
            return;
        }

        // 完全空的段落（`<w:p/>`）：也要留成块，前端用它撑出空行高度
        if content.is_empty() {
            out.push(Block::Paragraph(ParagraphBlock {
                style: style_id.as_deref().and_then(|id| self.styles.style_name(id)),
                style_id,
                align: ppr.align.clone(),
                indent_left_pt: ppr.indent_left_pt,
                indent_right_pt: ppr.indent_right_pt,
                indent_first_line_pt: ppr.first_line_pt,
                space_before_pt: ppr.space_before_pt,
                space_after_pt: ppr.space_after_pt,
                line_spacing: ppr.line.clone(),
                outline_level: ppr.outline_level,
                page_break_before,
                section_break,
                list,
                borders: ppr.borders,
                ..ParagraphBlock::default()
            }));
            return;
        }

        // 完全没有文字 run、只有图片/形状/占位对象的段落：各自成块
        // （前端按块渲染即可，不必再去解析段落内部；有换行 run 的空段落仍走下面的段落分支）
        if !has_visible_text && content.runs.is_empty() {
            for block in content.objects.drain(..) {
                out.push(block);
            }
            return;
        }

        out.push(Block::Paragraph(ParagraphBlock {
            style: style_id.as_deref().and_then(|id| self.styles.style_name(id)),
            style_id,
            runs: content.runs,
            text,
            align: ppr.align,
            indent_left_pt: ppr.indent_left_pt,
            indent_right_pt: ppr.indent_right_pt,
            indent_first_line_pt: ppr.first_line_pt,
            space_before_pt: ppr.space_before_pt,
            space_after_pt: ppr.space_after_pt,
            line_spacing: ppr.line,
            list,
            outline_level: ppr.outline_level,
            page_break_before,
            page_break: content.page_break,
            section_break,
            borders: ppr.borders,
        }));
        for block in content.objects {
            out.push(block);
        }
    }

    /// 行内节点（run / 超链接 / 修订包装 / 图形 …）
    fn walk_inline(
        &mut self,
        node: &XmlNode,
        base_rpr: &RunProps,
        char_style: Option<&str>,
        content: &mut ParagraphContent,
    ) {
        match node.local() {
            "r" => self.parse_run(node, base_rpr, char_style, content),
            "hyperlink" => {
                let style = node
                    .attr_local("rStyle")
                    .map(str::to_string)
                    .or_else(|| char_style.map(str::to_string));
                for child in node.children() {
                    self.walk_inline(child, base_rpr, style.as_deref(), content);
                }
            }
            "ins" | "smartTag" | "customXml" | "moveTo" => {
                for child in node.children() {
                    self.walk_inline(child, base_rpr, char_style, content);
                }
            }
            // 属性式域（`<w:fldSimple w:instr=" PAGE ">`）：里面的 run 就是缓存结果，
            // 没有 begin/separate，标记直接由 w:instr 决定
            "fldSimple" => {
                let saved = content.field.simple;
                let kind = node.attr_local("instr").and_then(field_kind_from_instruction);
                let before = content.runs.len();
                content.field.simple = kind;
                for child in node.children() {
                    self.walk_inline(child, base_rpr, char_style, content);
                }
                content.field.simple = saved;
                // 还没来得及算结果的属性式域（里面没有缓存文本）：同样补一个带标记的空 run
                if let Some(kind) = kind {
                    if content.runs.len() == before {
                        content.runs.push(Run {
                            text: String::new(),
                            field: Some(kind),
                            ..Run::default()
                        });
                    }
                }
            }
            // 域的分隔元素偶尔会直接挂在段落上（不包在 w:r 里）
            "fldChar" => {
                if let Some(kind) = content.field.on_fld_char(node) {
                    // 空结果的域：补一个带标记的空 run（详见 FieldState::on_fld_char）
                    content.runs.push(Run {
                        text: String::new(),
                        field: Some(kind),
                        ..Run::default()
                    });
                }
            }
            "instrText" => content.field.on_instruction(&node.texts()),
            "sdt" => {
                if let Some(inner) = node.child("sdtContent") {
                    for child in inner.children() {
                        self.walk_inline(child, base_rpr, char_style, content);
                    }
                }
            }
            // 修订痕迹：只显示最终态 —— w:del 里的文字不显示，w:ins 里的显示
            "del" | "moveFrom" | "delText" => {}
            "oMath" | "oMathPara" => {
                let plain = node.collect_text();
                let detail = if plain.trim().is_empty() {
                    "文档里有公式，Word / WPS 可正常显示".to_string()
                } else {
                    format!("公式内容（纯文本）：{}", plain.trim())
                };
                content
                    .objects
                    .push(unsupported("数学公式（OMML，暂不支持显示）", detail));
            }
            "AlternateContent" => self.walk_alternate(node, base_rpr, char_style, content),
            "drawing" => self.handle_drawing(node, content),
            "pict" => self.handle_pict(node, content),
            "object" => self.handle_object(node, content),
            "txbxContent" => {
                // 直接撞见文本框内容（例如别的包装节点里）：内容才是重点，几何用兜底值
                let look = self.text_box_look(node);
                match self.text_box(node, look) {
                    Some(text_box) => content.objects.push(Block::TextBox(text_box)),
                    None => {
                        let plain = node.collect_text();
                        let detail = if plain.trim().is_empty() {
                            "文本框里没有内容".to_string()
                        } else {
                            format!("文本框内容（纯文本）：{}", plain.trim())
                        };
                        content
                            .objects
                            .push(unsupported("文本框（暂不支持显示）", detail));
                    }
                }
            }
            // 段落/表格/单元格属性、锚点、书签：都不是可显示内容
            "pPr" | "rPr" | "tblPr" | "trPr" | "tcPr" | "sectPr" | "bookmarkStart"
            | "bookmarkEnd" | "proofErr" | "permStart" | "permEnd" | "commentRangeStart"
            | "commentRangeEnd" | "commentReference" | "footnoteReference" | "endnoteReference"
            | "lastRenderedPageBreak" | "annotationRef" => {}
            // 少数生成器会省掉 w:r 直接写 w:t —— 兜住，别丢字
            "t" => {
                let text = node.texts();
                if !text.is_empty() {
                    content.runs.push(Run {
                        text,
                        ..Run::default()
                    });
                }
            }
            _ => {
                // 不认识的包装节点：向下找可显示内容
                for child in node.children() {
                    self.walk_inline(child, base_rpr, char_style, content);
                }
            }
        }
    }

    /// `mc:AlternateContent`：优先 `mc:Choice`（新版表示），它没产出内容再用 `mc:Fallback`
    fn walk_alternate(
        &mut self,
        node: &XmlNode,
        base_rpr: &RunProps,
        char_style: Option<&str>,
        content: &mut ParagraphContent,
    ) {
        let before = (content.runs.len(), content.objects.len());
        if let Some(choice) = node.child("Choice") {
            for child in choice.children() {
                self.walk_inline(child, base_rpr, char_style, content);
            }
        }
        // Choice 有产出就不再看 Fallback（Word 里 Choice=DrawingML、Fallback=VML 是常态，
        // 两条都取会同一个图形产出两个块）
        let produced = content.runs.len() != before.0 || content.objects.len() != before.1;
        if !produced {
            if let Some(fallback) = node.child("Fallback") {
                for child in fallback.children() {
                    self.walk_inline(child, base_rpr, char_style, content);
                }
            }
        }
    }

    /// 一个 `w:r`：解析文本 + 其中的图形对象
    fn parse_run(
        &mut self,
        node: &XmlNode,
        base_rpr: &RunProps,
        char_style: Option<&str>,
        content: &mut ParagraphContent,
    ) {
        let mut rpr = node
            .child("rPr")
            .map(|node| parse_run_props(node, self.theme))
            .unwrap_or_default();
        let run_style = node
            .child("rPr")
            .and_then(|node| node.child("rStyle"))
            .and_then(|node| node.attr_local("val"))
            .map(str::to_string)
            .or_else(|| char_style.map(str::to_string));
        // 字符样式（含 basedOn）覆盖段落样式，但让位给 run 自己的直接格式
        if let Some(style) = run_style.as_deref() {
            rpr.apply_over(&self.styles.chain_rpr(Some(style)));
        }
        rpr.apply_over(base_rpr);

        let mut text = String::new();
        // 这个 run 的文本是不是某个"会变的域"的缓存结果（`PAGE` / `NUMPAGES`）
        let mut field: Option<FieldKind> = None;
        for child in node.children() {
            match child.local() {
                "t" => {
                    // 取**取文本那一刻**的域状态：`<w:r><w:fldChar begin/><w:instrText>..</w:instrText>
                    // <w:fldChar separate/><w:t>2</w:t><w:fldChar end/></w:r>` 这种"全塞一个 run"
                    // 的写法也要标对（run 结束时的状态早就出栈了）
                    let piece = child.texts();
                    if !piece.is_empty() {
                        content.field.note_text();
                        if field.is_none() {
                            field = content.field.result_kind();
                        }
                    }
                    text.push_str(&piece);
                }
                "tab" => text.push('\t'),
                "br" => {
                    // 分页符：独占一段时整段变成 PageBreak 块；与文字混排时保留换行
                    if child.attr_local("type") == Some("page") {
                        content.page_break = true;
                    }
                    text.push('\n');
                }
                "cr" => text.push('\n'),
                // 域的四个部分：指令文本（w:instrText）**永远不当正文输出**（Word 里它也不可见）
                "fldChar" => {
                    // 已识别、已算过、但结果区里一个字的缓存结果都没有 → 补一个带标记的空 run，
                    // 否则渲染器没有可替换的对象（新建文档的页脚会是一片空白）
                    if let Some(kind) = content.field.on_fld_char(child) {
                        content.runs.push(run_of(&rpr, String::new(), run_style.clone(), Some(kind)));
                    }
                }
                "instrText" => content.field.on_instruction(&child.texts()),
                "sym" => {
                    // 符号字符（`w:sym@w:char` 是十六进制码位）
                    if let Some(code) = child
                        .attr_local("char")
                        .and_then(|value| u32::from_str_radix(value.trim(), 16).ok())
                    {
                        if let Some(ch) = char::from_u32(code) {
                            text.push(ch);
                        }
                    }
                }
                "noBreakHyphen" => text.push('\u{2011}'),
                "softHyphen" => text.push('\u{00AD}'),
                "drawing" => self.handle_drawing(child, content),
                "pict" => self.handle_pict(child, content),
                "object" => self.handle_object(child, content),
                "txbxContent" => {
                    let look = self.text_box_look(child);
                    match self.text_box(child, look) {
                        Some(text_box) => content.objects.push(Block::TextBox(text_box)),
                        None => {
                            let plain = child.collect_text();
                            let detail = if plain.trim().is_empty() {
                                "文本框里没有内容".to_string()
                            } else {
                                format!("文本框内容（纯文本）：{}", plain.trim())
                            };
                            content
                                .objects
                                .push(unsupported("文本框（暂不支持显示）", detail));
                        }
                    }
                }
                "oMath" | "oMathPara" => {
                    let plain = child.collect_text();
                    let detail = if plain.trim().is_empty() {
                        "文档里有公式，Word / WPS 可正常显示".to_string()
                    } else {
                        format!("公式内容（纯文本）：{}", plain.trim())
                    };
                    content
                        .objects
                        .push(unsupported("数学公式（OMML，暂不支持显示）", detail));
                }
                "AlternateContent" => self.walk_alternate(child, base_rpr, char_style, content),
                _ => {}
            }
        }

        if !text.is_empty() {
            content
                .runs
                .push(run_of(&rpr, text, run_style.clone(), field));
        }
    }

    /// `w:drawing`：图片 / SmartArt / 图表 / 文本框 / 自选图形
    fn handle_drawing(&mut self, node: &XmlNode, content: &mut ParagraphContent) {
        // 先看 graphicData 的 URI：它决定这是图片还是图表/SmartArt
        if let Some(data) = node.find_descendant("graphicData") {
            if let Some(uri) = data.attr("uri") {
                if uri.contains("diagram") {
                    content.objects.push(unsupported(
                        "SmartArt 图形（暂不支持显示）",
                        "组织结构图 / 流程图等 SmartArt 需要 Word / WPS 渲染",
                    ));
                    return;
                }
                if uri.contains("chart") {
                    content.objects.push(unsupported(
                        "图表（暂不支持显示）",
                        "文档里的统计图表需要 Word / WPS 渲染",
                    ));
                    return;
                }
            }
        }

        if let Some(blip) = node.find_descendant("blip") {
            let rel_id = blip
                .attr_local("embed")
                .or_else(|| blip.attr_local("link"));
            if let Some(rel_id) = rel_id {
                match self.image_source(rel_id) {
                    Ok(media) => {
                        let (width, height) = node
                            .find_descendant("extent")
                            .map(|extent| {
                                (
                                    number(extent.attr_local("cx")).unwrap_or(0.0),
                                    number(extent.attr_local("cy")).unwrap_or(0.0),
                                )
                            })
                            .filter(|(cx, cy)| *cx > 0.0 && *cy > 0.0)
                            .map(|(cx, cy)| (emu_to_px(cx), emu_to_px(cy)))
                            .unwrap_or(FALLBACK_IMAGE_PX);
                        let doc_pr = node.find_descendant("docPr");
                        content.objects.push(Block::Image(ImageBlock {
                            media,
                            name: doc_pr.and_then(|node| node.attr("name")).map(str::to_string),
                            alt: doc_pr
                                .and_then(|node| node.attr("descr"))
                                .filter(|value| !value.trim().is_empty())
                                .map(str::to_string),
                            width_px: width,
                            height_px: height,
                        }));
                    }
                    // 关系缺失 / 指向非图片：说明白，别静默丢
                    Err(reason) => content
                        .objects
                        .push(unsupported("图片（无法显示）", reason)),
                }
                return;
            }
            content.objects.push(unsupported(
                "图形对象（暂不支持显示）",
                "这段内容用了本项目还不支持的图形表示，Word / WPS 可正常显示",
            ));
            return;
        }

        // 文本框：内容**结构化**成块（封面整页、表单签字框都是它）
        if let Some(content_node) = node.find_descendant("txbxContent") {
            let look = self.text_box_look(node);
            match self.text_box(content_node, look) {
                Some(text_box) => {
                    content.objects.push(Block::TextBox(text_box));
                }
                // 空文本框：仍然给占位（不静默丢）
                None => {
                    let plain = node.collect_text();
                    let detail = if plain.trim().is_empty() {
                        "文本框里没有内容".to_string()
                    } else {
                        format!("文本框内容（纯文本）：{}", plain.trim())
                    };
                    content
                        .objects
                        .push(unsupported("文本框（暂不支持显示）", detail));
                }
            }
            return;
        }

        // 自选图形（横线 / 方框）：能认出来就画，认不出来照旧给占位块
        if let Some(shape) = self.drawing_shape(node) {
            content.objects.push(Block::Shape(shape));
            return;
        }

        content.objects.push(unsupported(
            "图形对象（暂不支持显示）",
            "这段内容用了本项目还不支持的图形表示，Word / WPS 可正常显示",
        ));
    }

    /// DrawingML 自选图形（`wps:wsp` + `a:prstGeom`）→ [`ShapeBlock`]。
    /// 返回 `None` 表示"认不出来"，调用方会给占位块（**不丢内容**）：
    /// - 组合图形（`wpg:wgp`）：只渲染其中一部分会丢东西，整体走占位；
    /// - 没有 `prstGeom`（自由曲线 `a:custGeom`、连接符以外的复杂图形）；
    /// - 旋转 / 翻转过的形状（几何算不准，画歪不如不画）；
    /// - `prst` 不在支持表里（流程图、箭头、星形…）。
    fn drawing_shape(&self, node: &XmlNode) -> Option<ShapeBlock> {
        if node.has_descendant("wgp") {
            return None;
        }
        let wsp = node.find_descendant("wsp")?;
        let prst = wsp.find_descendant("prstGeom")?.attr("prst")?.trim();
        let shape = match prst {
            "line" => "line",
            // 直线连接符（Word 的"直线"连接符）在观感上就是一条直线
            "straightConnector1" => "line",
            "rect" => "rect",
            "roundRect" => "roundRect",
            "ellipse" => "ellipse",
            _ => return None,
        };
        if let Some(xfrm) = wsp.find_descendant("xfrm") {
            let rotated = xfrm
                .attr("rot")
                .and_then(|value| value.trim().parse::<f64>().ok())
                .is_some_and(|angle| angle != 0.0);
            if rotated || attr_is_on(xfrm, "flipH") || attr_is_on(xfrm, "flipV") {
                return None;
            }
        }

        // 尺寸优先用 `wp:extent`（Word 排版用的外框），退回形状自己的 `a:ext`
        let extent = |node: &XmlNode| -> Option<(f64, f64)> {
            let cx = number_f64(node.attr_local("cx").or_else(|| node.attr_local("w")))?;
            let cy = number_f64(node.attr_local("cy").or_else(|| node.attr_local("h")))?;
            Some((cx / EMU_PER_PT, cy / EMU_PER_PT))
        };
        let (mut width_pt, mut height_pt) = node
            .find_descendant("extent")
            .and_then(|node| extent(node))
            .or_else(|| wsp.find_descendant("ext").and_then(|node| extent(node)))
            .unwrap_or((0.0, 0.0));

        let (x_pt, y_pt) = wsp
            .find_descendant("off")
            .map(|off| {
                (
                    number_f64(off.attr_local("x")).unwrap_or(0.0) / EMU_PER_PT,
                    number_f64(off.attr_local("y")).unwrap_or(0.0) / EMU_PER_PT,
                )
            })
            .unwrap_or((0.0, 0.0));

        let sp_pr = wsp.find_descendant("spPr");
        let line = sp_pr.and_then(|sp| sp.child("ln"));
        let line_width_pt = line
            .and_then(|ln| number_f64(ln.attr("w")))
            .filter(|value| *value > 0.0)
            .map(|emu| emu / EMU_PER_PT);
        let line_color = line
            .and_then(|ln| ln.find_descendant("srgbClr").or_else(|| ln.find_descendant("sysClr")))
            .and_then(|color| color.attr("val").or_else(|| color.attr("lastClr")))
            .and_then(normalize_color_strict);
        let dash = line
            .and_then(|ln| ln.find_descendant("prstDash"))
            .and_then(|dash| dash.attr("val"))
            .filter(|value| !value.eq_ignore_ascii_case("solid"))
            .map(str::to_string);
        let fill_color = sp_pr
            .and_then(|sp| sp.child("solidFill"))
            .and_then(|fill| fill.find_descendant("srgbClr").or_else(|| fill.find_descendant("sysClr")))
            .and_then(|color| color.attr("val").or_else(|| color.attr("lastClr")))
            .and_then(normalize_color_strict);

        // 直线：某一维为 0 时用线宽（拿不到就 1pt）兜底，前端才好定位与命中
        if shape == "line" {
            let fallback = line_width_pt.unwrap_or(1.0);
            if height_pt <= 0.0 {
                height_pt = fallback;
            }
            if width_pt <= 0.0 {
                width_pt = fallback;
            }
        }

        Some(ShapeBlock {
            shape: shape.to_string(),
            x_pt,
            y_pt,
            width_pt,
            height_pt,
            line_width_pt,
            line_color,
            fill_color,
            dash,
            vertical: shape == "line" && height_pt > width_pt,
        })
    }

    /// `w:pict`（VML 老式图形）：图片 → 图片块；文本框 → 占位；线/框 → [`Block::Shape`]
    fn handle_pict(&mut self, node: &XmlNode, content: &mut ParagraphContent) {
        if let Some(image) = self.vml_image(node) {
            content.objects.push(Block::Image(image));
            return;
        }
        // VML 文本框（`v:shape` + `v:textbox`）：内容结构化，别只给占位
        if let Some(content_node) = node.find_descendant("txbxContent") {
            let look = self.text_box_look(node);
            match self.text_box(content_node, look) {
                Some(text_box) => content.objects.push(Block::TextBox(text_box)),
                None => {
                    let plain = node.collect_text();
                    let detail = if plain.trim().is_empty() {
                        "文本框里没有内容".to_string()
                    } else {
                        format!("文本框内容（纯文本）：{}", plain.trim())
                    };
                    content
                        .objects
                        .push(unsupported("文本框（暂不支持显示）", detail));
                }
            }
            return;
        }
        // 逐个**直接子形状**处理：一个 w:pict 里可能画了好几条线；
        // v:group 不拆（只画其中一部分会丢内容），交给占位块
        for child in node.children() {
            if let Some(shape) = self.vml_shape(child) {
                content.objects.push(Block::Shape(shape));
            } else {
                content.objects.push(unsupported(
                    "图形对象（VML，暂不支持显示）",
                    "这段内容用了老式 VML 图形（组合图形或本项目还不支持的形状），Word / WPS 可正常显示",
                ));
            }
        }
    }

    /// 文本框的几何 / 外观。取值优先级按"谁能给准就用谁"：
    /// 尺寸 `a:ext`（EMU）→ VML style 的 width/height → `wp:extent` → 兜底 200×40pt；
    /// 偏移 `a:off`（EMU）→ VML style 的 margin-left/top → 0；
    /// 填充与边框两套写法（DrawingML 的 `a:solidFill`/`a:ln`、VML 的 `fillcolor`/`strokecolor`）都认。
    fn text_box_look(&self, node: &XmlNode) -> TextBoxLook {
        let mut look = TextBoxLook::default();
        let sp_pr = node
            .find_descendant("wsp")
            .and_then(|wsp| wsp.find_descendant("spPr"));

        // ---- 尺寸 ----
        if let Some(size) = sp_pr
            .and_then(|sp| sp.find_descendant("xfrm"))
            .and_then(|xfrm| xfrm.child("ext"))
            .and_then(emu_extent)
        {
            (look.width_pt, look.height_pt) = size;
        } else if let Some(style) = vml_shape_style(node) {
            let (width, height, _) = vml_style_size(style);
            if let Some(width) = width.filter(|value| *value > 0.0) {
                look.width_pt = width;
            }
            if let Some(height) = height.filter(|value| *value > 0.0) {
                look.height_pt = height;
            }
        } else if let Some(size) = node.find_descendant("extent").and_then(emu_extent) {
            (look.width_pt, look.height_pt) = size;
        }

        // ---- 偏移 ----
        if let Some(off) = sp_pr
            .and_then(|sp| sp.find_descendant("xfrm"))
            .and_then(|xfrm| xfrm.child("off"))
        {
            look.x_pt = number_f64(off.attr_local("x")).unwrap_or(0.0) / EMU_PER_PT;
            look.y_pt = number_f64(off.attr_local("y")).unwrap_or(0.0) / EMU_PER_PT;
        } else if let Some(style) = vml_shape_style(node) {
            look.x_pt = vml_style_length(style, &["margin-left", "left"]).unwrap_or(0.0);
            look.y_pt = vml_style_length(style, &["margin-top", "top"]).unwrap_or(0.0);
        }

        // ---- 填充 / 边框 ----
        if let Some(sp_pr) = sp_pr {
            look.fill_color = drawing_fill(sp_pr);
            let (width, color) = drawing_line(sp_pr);
            look.border_width_pt = width;
            look.border_color = color;
        }
        if let Some(shape) = vml_shape_node(node) {
            let filled = !matches!(
                shape.attr_local("filled"),
                Some("f") | Some("false") | Some("0")
            );
            if look.fill_color.is_none() && filled {
                look.fill_color = shape.attr_local("fillcolor").and_then(normalize_color_strict);
            }
            if look.border_color.is_none() {
                look.border_color = shape.attr_local("strokecolor").and_then(normalize_color_strict);
            }
            if look.border_width_pt.is_none() {
                look.border_width_pt = shape
                    .attr_local("strokeweight")
                    .and_then(parse_style_length)
                    .map(|(value, factor)| value * factor);
            }
        }

        look.wrap = drawing_wrap(node)
            .or_else(|| vml_shape_style(node).and_then(vml_wrap))
            .unwrap_or_else(|| "none".to_string());
        look
    }

    /// `w:txbxContent` → [`TextBoxBlock`]：内容走**同一套**块解析（段落样式层叠、编号、嵌套表格…）。
    /// 返回 `None` 表示里面没有任何块（空文本框），调用方给占位块。
    fn text_box(&mut self, content_node: &XmlNode, look: TextBoxLook) -> Option<TextBoxBlock> {
        let mut blocks = Vec::new();
        // 文本框是独立的"故事"（story）：它自己的列表从 1 开始数，
        // 也不会把正文的编号计数器往后推 —— 所以计数器进出都还原。
        let saved = std::mem::take(&mut self.numbering_state.counters);
        self.collect_blocks(content_node, &mut blocks);
        self.numbering_state.counters = saved;
        if blocks.is_empty() {
            return None;
        }
        Some(TextBoxBlock {
            blocks,
            x_pt: look.x_pt,
            y_pt: look.y_pt,
            width_pt: look.width_pt,
            height_pt: look.height_pt,
            fill_color: look.fill_color,
            border_color: look.border_color,
            border_width_pt: look.border_width_pt,
            wrap: look.wrap,
        })
    }

    /// VML 形状（`v:line` / `v:rect` / `v:roundrect` / `v:oval` / `v:hr`）→ [`ShapeBlock`]。
    /// 认不出来（`v:group`、`v:shape`、`v:polyline`…）返回 `None`。
    fn vml_shape(&self, node: &XmlNode) -> Option<ShapeBlock> {
        let (shape, is_line) = match node.local() {
            "line" => ("line", true),
            "rect" => ("rect", false),
            "roundrect" => ("roundRect", false),
            "oval" => ("ellipse", false),
            "hr" => {
                // Word 的「水平线」：没有尺寸信息，就是整栏宽的一条细线
                return Some(ShapeBlock {
                    shape: "line".to_string(),
                    width_pt: self.current_width_pt,
                    height_pt: 1.0,
                    vertical: false,
                    ..ShapeBlock::default()
                });
            }
            _ => return None,
        };

        let style = node.attr("style").unwrap_or_default();
        let (width, height, unit_pt) = vml_style_size(style);
        let mut width_pt = width.unwrap_or(0.0);
        let mut height_pt = height.unwrap_or(0.0);
        // 偏移：VML 写在 style 的 margin-left / margin-top（或 left / top）里
        let x_pt = vml_style_length(style, &["margin-left", "left"]).unwrap_or(0.0);
        let y_pt = vml_style_length(style, &["margin-top", "top"]).unwrap_or(0.0);

        // 直线：端点（`from` / `to`）用来判方向；坐标空间见 parse_vml_point
        let mut vertical = false;
        if is_line {
            let from = node.attr("from").and_then(parse_vml_point);
            let to = node.attr("to").and_then(parse_vml_point);
            if let (Some(from), Some(to)) = (from, to) {
                let coordsize = node
                    .attr_local("coordsize")
                    .and_then(|value| parse_vml_point(value).map(|point| (point.0, point.1)));
                let begin = resolve_vml_point(from, coordsize, (width_pt, height_pt), unit_pt);
                let end = resolve_vml_point(to, coordsize, (width_pt, height_pt), unit_pt);
                let dx = (end.0 - begin.0).abs();
                let dy = (end.1 - begin.1).abs();
                // 样式没给尺寸时，用两端点的包围盒
                if width_pt <= 0.0 && dx > 0.0 {
                    width_pt = dx;
                }
                if height_pt <= 0.0 && dy > 0.0 {
                    height_pt = dy;
                }
                vertical = dy > dx;
            }
        }

        let stroke_width = node
            .attr_local("strokeweight")
            .and_then(parse_style_length)
            .map(|(value, factor)| value * factor);
        let stroke_color = node
            .attr_local("strokecolor")
            .and_then(normalize_color_strict);
        let filled = !matches!(node.attr_local("filled"), Some("f") | Some("false") | Some("0"));
        let fill_color = node
            .attr_local("fillcolor")
            .and_then(normalize_color_strict)
            .filter(|_| filled);
        let dash = node
            .attr_local("dashstyle")
            .filter(|value| !value.eq_ignore_ascii_case("solid"))
            .map(|value| normalize_vml_dash(value).to_string());

        if is_line {
            let fallback = stroke_width.unwrap_or(1.0);
            if width_pt <= 0.0 {
                width_pt = fallback;
            }
            if height_pt <= 0.0 {
                height_pt = fallback;
            }
            vertical = vertical || height_pt > width_pt;
        }

        Some(ShapeBlock {
            shape: shape.to_string(),
            x_pt,
            y_pt,
            width_pt,
            height_pt,
            line_width_pt: stroke_width,
            line_color: stroke_color,
            fill_color,
            dash,
            vertical,
        })
    }

    /// `w:object`（OLE 嵌入对象，如公式编辑器 3.0、Excel 表格对象）。
    /// 这类对象在文档里通常自带一张**预览图**，把预览图显示出来并附占位说明，
    /// 比只给一个空占位卡片有用得多。
    fn handle_object(&mut self, node: &XmlNode, content: &mut ParagraphContent) {
        let mut previewed = false;
        if let Some(image) = self.vml_image(node) {
            content.objects.push(Block::Image(image));
            previewed = true;
        }
        let prog_id = node
            .find_descendant("OLEObject")
            .and_then(|ole| ole.attr("ProgID"))
            .map(|prog| format!("（{prog}）"))
            .unwrap_or_default();
        content.objects.push(unsupported(
            "OLE 嵌入对象（暂不支持显示）",
            if previewed {
                format!("上面显示的是文档里的预览图{prog_id}；打开/编辑实际对象请用 Word / WPS")
            } else {
                format!("文档里嵌入了对象{prog_id}，需要 Word / WPS 打开查看")
            },
        ));
    }

    /// VML 图片（`v:imagedata`），尺寸从 `v:shape@style` 的 width/height 里读
    fn vml_image(&self, node: &XmlNode) -> Option<ImageBlock> {
        let imagedata = node.find_descendant("imagedata")?;
        let rel_id = imagedata
            .attr_local("id")
            .or_else(|| imagedata.attr_local("relid"))
            .or_else(|| imagedata.attr_local("embed"))?;
        let media = self.image_source(rel_id).ok()?;
        let (width, height) = node
            .find_descendant("shape")
            .and_then(|shape| shape.attr("style"))
            .map(parse_vml_size)
            .unwrap_or(FALLBACK_IMAGE_PX);
        Some(ImageBlock {
            media,
            name: node
                .find_descendant("shape")
                .and_then(|shape| shape.attr("id"))
                .map(str::to_string),
            alt: imagedata.attr_local("title").map(str::to_string),
            width_px: width,
            height_px: height,
        })
    }

    /// 关系 id → 图片块要用的媒体地址。
    /// 包内部件给包内路径（前端走 asset 协议）；外部链接给 URL 原样带回；
    /// 关系缺失、或关系指向的不是图片，返回中文原因（调用方转成占位块）。
    fn image_source(&self, rel_id: &str) -> Result<String, String> {
        let Some(rel) = self.rels.get(rel_id) else {
            return Err(format!(
                "文档引用了关系 {rel_id}，但 document.xml.rels 里找不到它，图片部件可能已丢失"
            ));
        };
        if !rel.external && !rel.kind.is_empty() && rel.kind != "image" {
            return Err(format!(
                "关系 {rel_id} 指向的不是图片（{}：{}）",
                rel.kind, rel.target
            ));
        }
        Ok(rel.target.clone())
    }

    /// 一个 `w:tbl`
    fn build_table(&mut self, node: &XmlNode) -> Block {
        let pr = node.child("tblPr");
        let style_id = pr
            .and_then(|pr| pr.child("tblStyle"))
            .and_then(|node| node.attr_local("val"))
            .map(str::to_string);
        let align = pr
            .and_then(|pr| pr.child("jc"))
            .and_then(|node| node.attr_local("val"))
            .map(str::to_string);

        // 表格总宽：只认 dxa（twip）；pct/auto 交给前端按列宽算
        let width_px = pr
            .and_then(|pr| pr.child("tblW"))
            .filter(|w| w.attr_local("type").map(|kind| kind == "dxa").unwrap_or(true))
            .and_then(|w| number(w.attr_local("w")))
            .filter(|value| *value > 0.0)
            .map(twips_to_px);

        // 表格级边框：作为单元格没写边框时的兜底
        let table_borders = pr
            .and_then(|pr| pr.child("tblBorders"))
            .map(parse_borders)
            .unwrap_or_default();

        let columns: Vec<f32> = node
            .child("tblGrid")
            .map(|grid| {
                grid.children_named("gridCol")
                    .filter_map(|col| number(col.attr_local("w")))
                    .filter(|value| *value > 0.0)
                    .map(twips_to_px)
                    .collect()
            })
            .unwrap_or_default();

        let mut rows: Vec<TableRow> = Vec::new();
        let table_grid: Vec<f32> = columns.clone();
        for row in node.children_named("tr").chain(
            // 少数文档把 w:tr 包在 w:sdt 里
            node.children_named("sdt")
                .filter_map(|sdt| sdt.child("sdtContent"))
                .flat_map(|content| content.children_named("tr")),
        ) {
            rows.push(self.build_row(row, &table_borders, &table_grid));
        }

        let borders = table_borders.iter().any(|side| *side == Some(true))
            || rows.iter().any(|row| {
                row.cells
                    .iter()
                    .any(|cell| cell.borders.top || cell.borders.left || cell.borders.bottom || cell.borders.right)
            });

        Block::Table(TableBlock {
            rows,
            columns,
            width_px,
            align,
            borders,
            style_id,
        })
    }

    fn build_row(
        &mut self,
        node: &XmlNode,
        table_borders: &[Option<bool>; 4],
        grid: &[f32],
    ) -> TableRow {
        let mut row = TableRow::default();
        if let Some(tr_pr) = node.child("trPr") {
            row.header = tr_pr.child("tblHeader").is_some();
            row.height_px = tr_pr
                .child("trHeight")
                .and_then(|height| number(height.attr_local("val")))
                .filter(|value| *value > 0.0)
                .map(twips_to_px);
        }

        let mut column_index = 0usize;
        for cell in node.children_named("tc") {
            let tc_pr = cell.child("tcPr");
            let mut block = TableCell::default();

            block.grid_span = tc_pr
                .and_then(|pr| pr.child("gridSpan"))
                .and_then(|span| integer(span.attr_local("val")))
                .unwrap_or(1)
                .clamp(1, 100) as usize;

            block.v_merge = match tc_pr.and_then(|pr| pr.child("vMerge")) {
                None => VMerge::None,
                Some(merge) => match merge.attr_local("val") {
                    Some("restart") => VMerge::Restart,
                    _ => VMerge::Continue,
                },
            };

            if let Some(width) = tc_pr.and_then(|pr| pr.child("tcW")) {
                let is_dxa = width.attr_local("type").map(|kind| kind == "dxa").unwrap_or(true);
                if is_dxa {
                    block.width_px = number(width.attr_local("w"))
                        .filter(|value| *value > 0.0)
                        .map(twips_to_px);
                }
            }
            // tcW 常写成百分比，这时退而用表格网格里的列宽（前两列通常够用）
            if block.width_px.is_none() && column_index < grid.len() {
                block.width_px = Some(grid[column_index]);
            }

            block.shading = tc_pr
                .and_then(|pr| pr.child("shd"))
                .and_then(|shd| shd.attr_local("fill"))
                .and_then(normalize_color);
            block.v_align = tc_pr
                .and_then(|pr| pr.child("vAlign"))
                .and_then(|node| node.attr_local("val"))
                .map(str::to_string);

            // 单元格边框：自己写了的边用自己的，没写的边沿用表格级（逐边合并）
            let own = tc_pr
                .and_then(|pr| pr.child("tcBorders"))
                .map(parse_borders);
            let borders = merge_borders(own, *table_borders);
            block.borders = CellBorders {
                top: borders[0],
                left: borders[1],
                bottom: borders[2],
                right: borders[3],
            };

            // 单元格里的块（多个段落 / 嵌套表格）
            let mut blocks = Vec::new();
            self.collect_blocks(cell, &mut blocks);
            block.text = blocks
                .iter()
                .map(block_plain_text)
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>()
                .join("\n");
            block.blocks = blocks;

            column_index += block.grid_span;
            row.cells.push(block);
        }
        row
    }
}

/// `w:tblBorders` / `w:tcBorders` → [上, 左, 下, 右]。
/// `Some(true/false)` = 这一侧明确写了（可见 / `nil`、`none` 明确不可见），
/// `None` = 这一侧没写 —— 单元格要按"边"去沿用表格级设置，不能整块覆盖。
fn parse_borders(node: &XmlNode) -> [Option<bool>; 4] {
    let side = |name: &str| -> Option<bool> {
        let border = node.child(name)?;
        Some(!matches!(border.attr_local("val"), Some("nil") | Some("none")))
    };
    [side("top"), side("left"), side("bottom"), side("right")]
}

/// 单元格四边：自己写了的边用自己的，没写的边沿用表格级
fn merge_borders(cell: Option<[Option<bool>; 4]>, table: [Option<bool>; 4]) -> [bool; 4] {
    let mut out = [false; 4];
    for (index, slot) in out.iter_mut().enumerate() {
        *slot = cell
            .and_then(|sides| sides[index])
            .or(table[index])
            .unwrap_or(false);
    }
    out
}

/// VML `style="width:72pt;height:36pt"` → px（认不出就给兜底尺寸，绝不给 0）。
/// 注意这里要的是 **px**（图片块用 px），而 [`vml_style_size`] 给的是 pt，所以 ×4/3。
fn parse_vml_size(style: &str) -> (f32, f32) {
    let (width, height, _) = vml_style_size(style);
    (
        width
            .map(|pt| (pt * 4.0 / 3.0) as f32)
            .unwrap_or(FALLBACK_IMAGE_PX.0),
        height
            .map(|pt| (pt * 4.0 / 3.0) as f32)
            .unwrap_or(FALLBACK_IMAGE_PX.1),
    )
}

/// VML / CSS style 里的长度 → `(数值, 换算成 pt 的系数)`。
/// 无单位的长度按 CSS 约定当 px（×0.75）——Word 写的 VML 基本都带单位，这条是兜底。
fn parse_style_length(value: &str) -> Option<(f64, f64)> {
    let value = value.trim();
    for (suffix, factor) in [
        ("pt", 1.0),
        ("px", 0.75),
        ("in", 72.0),
        ("cm", 72.0 / 2.54),
        ("mm", 72.0 / 25.4),
        ("pc", 12.0),
        ("em", 12.0),
    ] {
        if let Some(number) = value.strip_suffix(suffix) {
            return number
                .trim()
                .parse::<f64>()
                .ok()
                .map(|number| (number, factor));
        }
    }
    value.parse::<f64>().ok().map(|number| (number, 0.75))
}

/// VML `style` 里的 width / height → `(宽 pt, 高 pt, 该 style 的"单位系数")`。
/// 第三个值是给**无单位**的 `from`/`to` 坐标用的：VML 规范里没写 `coordsize` 时，
/// 坐标空间就是形状自身的尺寸，单位与 style 一致（Word 写 `width:400pt` + `to="400,0"`）。
fn vml_style_size(style: &str) -> (Option<f64>, Option<f64>, f64) {
    let mut width = None;
    let mut height = None;
    let mut unit_pt = 1.0;
    for part in style.split(';') {
        let Some((key, value)) = part.split_once(':') else {
            continue;
        };
        let Some((number, factor)) = parse_style_length(value) else {
            continue;
        };
        match key.trim().to_ascii_lowercase().as_str() {
            "width" => {
                width = Some(number * factor);
                unit_pt = factor;
            }
            "height" => {
                height = Some(number * factor);
                unit_pt = factor;
            }
            _ => {}
        }
    }
    (width, height, unit_pt)
}

/// 从 VML `style` 里取第一个命中的长度属性（pt）：`margin-left` / `left` / `margin-top` / `top`
fn vml_style_length(style: &str, keys: &[&str]) -> Option<f64> {
    for part in style.split(';') {
        let Some((key, value)) = part.split_once(':') else {
            continue;
        };
        let key = key.trim().to_ascii_lowercase();
        if keys.iter().any(|wanted| *wanted == key) {
            if let Some((number, factor)) = parse_style_length(value) {
                return Some(number * factor);
            }
        }
    }
    None
}

/// VML 点（`from="0,0"` / `to="400pt,0"`）→ `(x, y, 是否带单位)`
fn parse_vml_point(value: &str) -> Option<(f64, f64, bool)> {
    let mut parts = value.split(',');
    let x = parts.next()?.trim();
    let y = parts.next().unwrap_or("0").trim();
    let (x_number, x_factor) = parse_style_length(x)?;
    let (y_number, y_factor) = parse_style_length(y)?;
    // 判断"带没带单位"：带单位时 parse_style_length 走的不是无单位分支（系数 ≠ 0.75 或缺后缀）
    let has_unit = |raw: &str| {
        let raw = raw.trim();
        raw.ends_with("pt") || raw.ends_with("px") || raw.ends_with("in") || raw.ends_with("cm")
            || raw.ends_with("mm") || raw.ends_with("pc") || raw.ends_with("em")
    };
    let unit = has_unit(x) || has_unit(y);
    let (x_pt, y_pt) = if unit {
        (x_number * x_factor, y_number * y_factor)
    } else {
        (x_number, y_number)
    };
    Some((x_pt, y_pt, unit))
}

/// VML 端点 → pt。三种坐标空间，按优先级：
/// 1. 端点自带单位 → 已经是 pt，直接用；
/// 2. 有 `coordsize`（`coordsize="1000,1000"`）→ 按它映射到形状的 style 尺寸；
/// 3. 都没有 → 坐标与 style 同单位（Word 的常见写法）。
fn resolve_vml_point(
    point: (f64, f64, bool),
    coordsize: Option<(f64, f64)>,
    box_pt: (f64, f64),
    unit_pt: f64,
) -> (f64, f64) {
    if point.2 {
        return (point.0, point.1);
    }
    match coordsize {
        Some((cx, cy)) if cx > 0.0 && cy > 0.0 => (
            point.0 / cx * box_pt.0.max(1.0),
            point.1 / cy * box_pt.1.max(1.0),
        ),
        _ => (point.0 * unit_pt, point.1 * unit_pt),
    }
}

/// VML 的 `dashstyle` 名字归一到与 DrawingML `a:prstDash` 同名的那一套
fn normalize_vml_dash(value: &str) -> &'static str {
    match value.trim().to_ascii_lowercase().as_str() {
        "shortdot" | "dot" => "dot",
        "shortdash" | "dash" => "dash",
        "dashdot" | "shortdashdot" => "dashDot",
        "longdash" | "dashdotdot" | "longdashdot" => "lgDash",
        "sysdot" => "sysDot",
        "sysdash" => "sysDash",
        _ => "dash",
    }
}

/// XML 属性形式的布尔（`val="1"` / `"true"` / `"on"` 为真；缺省为假）
fn attr_is_on(node: &XmlNode, name: &str) -> bool {
    match node.attr_local(name) {
        None => false,
        Some(value) => !matches!(value.trim(), "0" | "false" | "off" | "none"),
    }
}

/// 文本框拿不到尺寸时的兜底（pt）—— 宁可给个能看见的框，也不要尺寸为 0 让内容消失
const TEXTBOX_FALLBACK_PT: (f64, f64) = (200.0, 40.0);

/// 文本框的几何与外观（[`DocBuilder::text_box_look`] 的产物）
struct TextBoxLook {
    x_pt: f64,
    y_pt: f64,
    width_pt: f64,
    height_pt: f64,
    fill_color: Option<String>,
    border_color: Option<String>,
    border_width_pt: Option<f64>,
    wrap: String,
}

impl Default for TextBoxLook {
    fn default() -> Self {
        TextBoxLook {
            x_pt: 0.0,
            y_pt: 0.0,
            width_pt: TEXTBOX_FALLBACK_PT.0,
            height_pt: TEXTBOX_FALLBACK_PT.1,
            fill_color: None,
            border_color: None,
            border_width_pt: None,
            wrap: "none".to_string(),
        }
    }
}

/// `a:ext` / `wp:extent`（EMU）→ `(宽 pt, 高 pt)`；缺 cx/cy 或非数字返回 `None`
fn emu_extent(node: &XmlNode) -> Option<(f64, f64)> {
    let cx = number_f64(node.attr_local("cx").or_else(|| node.attr_local("w")))?;
    let cy = number_f64(node.attr_local("cy").or_else(|| node.attr_local("h")))?;
    Some((cx / EMU_PER_PT, cy / EMU_PER_PT))
}

/// `a:sfrgba`…——DrawingML 填充色。只看 `spPr` 的**直接子** `a:solidFill`：
/// `a:ln/a:solidFill` 是描边色，别混进来。
fn drawing_fill(sp_pr: &XmlNode) -> Option<String> {
    sp_pr
        .child("solidFill")
        .and_then(|fill| fill.find_descendant("srgbClr").or_else(|| fill.find_descendant("sysClr")))
        .and_then(|color| color.attr("val").or_else(|| color.attr("lastClr")))
        .and_then(normalize_color_strict)
}

/// DrawingML 描边：`spPr/a:ln@w`（EMU → pt）与 `a:ln` 里的颜色
fn drawing_line(sp_pr: &XmlNode) -> (Option<f64>, Option<String>) {
    let Some(line) = sp_pr.child("ln") else {
        return (None, None);
    };
    let width = number_f64(line.attr("w"))
        .filter(|value| *value > 0.0)
        .map(|emu| emu / EMU_PER_PT);
    let color = line
        .find_descendant("srgbClr")
        .or_else(|| line.find_descendant("sysClr"))
        .and_then(|color| color.attr("val").or_else(|| color.attr("lastClr")))
        .and_then(normalize_color_strict);
    (width, color)
}

/// DrawingML 环绕方式：`wp:anchor` 里的 `wp:wrapXxx` 子节点 → `square` / `none` / `tight` /
/// `through` / `topAndBottom`；内联（没有 `wp:anchor`）返回 `None`（由调用方定 "none"）。
fn drawing_wrap(node: &XmlNode) -> Option<String> {
    let anchor = node.child("anchor")?;
    for child in anchor.children() {
        let Some(kind) = child.local().strip_prefix("wrap") else {
            continue;
        };
        let mut chars = kind.chars();
        let first = chars.next().unwrap_or('n').to_ascii_lowercase();
        return Some(format!("{first}{}", chars.as_str()));
    }
    None
}

/// VML 形状节点：`w:pict` 的直接子形状，或（`w:object` 那种）后代里的第一个
fn vml_shape_node(node: &XmlNode) -> Option<&XmlNode> {
    if matches!(
        node.local(),
        "shape" | "rect" | "roundrect" | "oval" | "line" | "group" | "polyline"
    ) {
        return Some(node);
    }
    node.children().find(|child| {
        matches!(
            child.local(),
            "shape" | "rect" | "roundrect" | "oval" | "line" | "group" | "polyline"
        )
    })
}

/// VML 形状的 `style` 字符串（`w:pict` 节点或形状节点都能给）
fn vml_shape_style(node: &XmlNode) -> Option<&str> {
    vml_shape_node(node)?.attr("style")
}

/// VML 的 `mso-wrap-style` → 与 DrawingML 同一套环绕名字
fn vml_wrap(style: &str) -> Option<String> {
    let value = vml_style_named(style, "mso-wrap-style")?;
    Some(match value.to_ascii_lowercase().as_str() {
        "square" => "square".to_string(),
        "tight" => "tight".to_string(),
        "through" => "through".to_string(),
        "topandbottom" => "topAndBottom".to_string(),
        _ => "none".to_string(),
    })
}

/// 从 VML `style` 里取一个**非长度**的具名属性值（`mso-wrap-style:square`）
fn vml_style_named<'a>(style: &'a str, name: &str) -> Option<&'a str> {
    for part in style.split(';') {
        let Some((key, value)) = part.split_once(':') else {
            continue;
        };
        if key.trim().eq_ignore_ascii_case(name) {
            return Some(value.trim());
        }
    }
    None
}

/// 占位块构造（`label` 是给用户看的中文说明）
fn unsupported(label: &str, detail: impl Into<String>) -> Block {
    Block::Unsupported {
        label: label.to_string(),
        detail: detail.into(),
    }
}

/// 块的纯文本（查找索引 / 表格单元格文本用）
fn block_plain_text(block: &Block) -> String {
    match block {
        Block::Paragraph(paragraph) => paragraph.text.clone(),
        Block::Table(table) => table
            .rows
            .iter()
            .map(|row| {
                row.cells
                    .iter()
                    .map(|cell| cell.text.clone())
                    .collect::<Vec<_>>()
                    .join("\t")
            })
            .collect::<Vec<_>>()
            .join("\n"),
        Block::Image(image) => image.name.clone().unwrap_or_default(),
        // 形状没有文字（查找"直线/方框"没有意义，别污染搜索结果）
        Block::Shape(_) => String::new(),
        // 文本框里的文字是**正文内容**（封面标题、表单字段…），要能被搜到
        Block::TextBox(text_box) => text_box
            .blocks
            .iter()
            .map(block_plain_text)
            .filter(|text| !text.is_empty())
            .collect::<Vec<_>>()
            .join("\n"),
        Block::PageBreak => String::new(),
        Block::Unsupported { label, detail } => {
            if detail.is_empty() {
                label.clone()
            } else {
                format!("{label} {detail}")
            }
        }
    }
}

/* ================================================================================== */
/* 七、文档解析入口 + 内存缓存                                                        */
/* ================================================================================== */

/// 解析结果。`document_blocks` / `document_find` 共用，靠缓存避免虚拟滚动时反复解析。
struct ParsedDocument {
    blocks: Vec<Block>,
    /// 查找索引：顶层块下标 → 该块的可检索文本（表格用整表文本，命中后定位到表格这一块）
    search_index: Vec<(usize, String)>,
    /// 文档里 `w:p` 元素总数（含表格单元格内的）——用于与 `document_info` 的段落统计对账
    /// （解析完整性的自检指标；非测试构建下没人读它）
    #[cfg_attr(not(test), allow(dead_code))]
    paragraph_count: usize,
    encrypted: bool,
    /// 页面几何（`w:sectPr`；缺失为 None）。解析一次缓存起来，每次取窗口都带上。
    page: Option<PageGeometry>,
    /// 文档自己的页眉（`w:headerReference`，默认页眉）
    header: Option<Vec<Block>>,
    /// 文档自己的页脚
    footer: Option<Vec<Block>>,
}

/// 缓存最近打开的 2 个文档（只读查看，缓存是安全的；解析大文档要几十毫秒，虚拟滚动会反复要）
const CACHE_DOCUMENTS: usize = 2;
/// 单次最多返回的块数（防止前端误传超大 count 把整篇文档序列化过去）
const MAX_BLOCK_WINDOW: usize = 2000;
/// 查找命中上限（与表格模块一致）
const MAX_FIND_HITS: usize = 500;
/// 命中上下文片段的左右留白（字符数）
const EXCERPT_PAD: usize = 30;
/// 单张图片的大小上限（32MB）：超过就明确拒绝，别把几十 MB 解压进内存再 base64 传前端
const MAX_MEDIA_BYTES: usize = 32 * 1024 * 1024;

/// A4 竖版纸张尺寸（pt）：`w:sectPr` 缺失或 `w:pgSz` 畸形时的兜底
const A4_WIDTH_PT: f64 = 595.3;
const A4_HEIGHT_PT: f64 = 841.9;
/// A4 默认页边距（pt）= 2.54cm：`w:pgMar` 缺失或畸形时的兜底
const A4_MARGIN_PT: f64 = 72.0;

struct CacheEntry {
    /// 路径原样（用户视角的 key）
    key: String,
    /// 大小 + 修改时间：文件变了就重新解析
    stamp: String,
    doc: Arc<ParsedDocument>,
}

static DOCUMENT_CACHE: OnceLock<Mutex<Vec<CacheEntry>>> = OnceLock::new();

fn cache() -> &'static Mutex<Vec<CacheEntry>> {
    DOCUMENT_CACHE.get_or_init(|| Mutex::new(Vec::new()))
}

fn cache_stamp(path: &Path) -> Option<String> {
    let meta = std::fs::metadata(path).ok()?;
    let modified = meta
        .modified()
        .ok()
        .and_then(|time| time.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|duration| duration.as_nanos())
        .unwrap_or(0);
    Some(format!("{}:{modified}", meta.len()))
}

/// 取（必要时解析）文档；命中缓存时只做一次元数据检查
fn load_document(path: &str) -> Result<Arc<ParsedDocument>, String> {
    let file = Path::new(path);
    let stamp = cache_stamp(file);
    if let Some(stamp) = stamp.as_deref() {
        let entries = cache().lock().unwrap_or_else(|error| error.into_inner());
        if let Some(entry) = entries
            .iter()
            .find(|entry| entry.key == path && entry.stamp == stamp)
        {
            return Ok(entry.doc.clone());
        }
    }

    let (bytes, encrypted) = read_document_bytes(file)?;
    let doc = Arc::new(parse_document(&bytes, encrypted)?);

    if let Some(stamp) = stamp {
        let mut entries = cache().lock().unwrap_or_else(|error| error.into_inner());
        // 同一个文件的新版本：先把旧的挤掉
        entries.retain(|entry| entry.key != path);
        entries.insert(
            0,
            CacheEntry {
                key: path.to_string(),
                stamp,
                doc: doc.clone(),
            },
        );
        entries.truncate(CACHE_DOCUMENTS);
    }
    Ok(doc)
}

/// 把整个 docx 解析成块模型。**任何路径上的"缺东西"都降级处理**，不 panic。
fn parse_document(bytes: &[u8], encrypted: bool) -> Result<ParsedDocument, String> {
    let mut zip = open_zip(bytes)?;

    let document_xml = read_part_lossy(&mut zip, "word/document.xml")
        .ok_or_else(|| "包内没有 word/document.xml，可能不是 Word 文档".to_string())?;
    // 下面这些都是"有则用、没有就用默认"，缺失不影响正文显示
    let styles_xml = read_part_lossy(&mut zip, "word/styles.xml").unwrap_or_default();
    let numbering_xml = read_part_lossy(&mut zip, "word/numbering.xml").unwrap_or_default();
    let rels_xml = read_part_lossy(&mut zip, "word/_rels/document.xml.rels").unwrap_or_default();
    let theme_xml = read_part_lossy(&mut zip, "word/theme/theme1.xml").unwrap_or_default();

    let root = parse_xml(&document_xml)
        .ok_or_else(|| "word/document.xml 解析不出内容（文件可能已损坏）".to_string())?;
    let body = find_body(&root)
        .ok_or_else(|| "这个 .docx 里没有正文（缺少 w:body），文件可能已损坏".to_string())?;

    let theme = Theme::parse(&theme_xml);
    let styles = StyleSheet::parse(&styles_xml, &theme);
    let numbering = Numbering::parse(&numbering_xml);
    let rels = parse_rels(&rels_xml);
    // 页面几何要先算：`v:hr`（整栏宽的水平线）之类的形状需要栏宽，块构建时就要用
    let section = find_section_properties(body);
    let page = section.map(parse_page_geometry);

    let mut builder = DocBuilder::new(&styles, &numbering, &theme, &rels, page.as_ref());
    let mut blocks = Vec::new();
    builder.collect_blocks(body, &mut blocks);

    let search_index = blocks
        .iter()
        .enumerate()
        .map(|(index, block)| (index, block_plain_text(block)))
        .filter(|(_, text)| !text.trim().is_empty())
        .collect();

    // 文档自己的页眉页脚（`w:headerReference` / `w:footerReference`）：
    // 各自是一个独立"故事"，里面的列表编号不该扰动正文计数器。
    // 故事里"什么都没有"（全是空段落，例如只有一个占位的空页脚）时按 `None` 返回，
    // 免得前端为一片空白留出页眉页脚的高度。
    let header = read_story(&mut zip, &rels, body, "headerReference", "hdr")
        .map(|container| builder.parse_story(&container))
        .filter(|blocks| !story_is_empty(blocks));
    let footer = read_story(&mut zip, &rels, body, "footerReference", "ftr")
        .map(|container| builder.parse_story(&container))
        .filter(|blocks| !story_is_empty(blocks));

    Ok(ParsedDocument {
        blocks,
        search_index,
        paragraph_count: builder.paragraph_count,
        encrypted,
        page,
        header,
        footer,
    })
}

/// 解析一段独立"故事"（页眉 / 页脚 / 文本框都能用）：
/// 先把编号计数器存档，解析完再还原 —— 故事里的编号与正文互不影响。
fn read_story(
    zip: &mut ZipArchive<Cursor<&[u8]>>,
    rels: &HashMap<String, RelInfo>,
    body: &XmlNode,
    reference: &str,
    root_local: &str,
) -> Option<XmlNode> {
    let target = story_target(body, rels, reference)?;
    let xml = read_part_lossy(zip, &target)?;
    parse_story_xml(&xml, root_local)
}

/// 按**文档顺序**列出所有分节设置：先段落级的（分节符，各自代表它前面那一节），
/// 最后是 `w:body/w:sectPr`（Word 把最后一节的设置放在这儿）。
fn sections_in_order(body: &XmlNode) -> Vec<&XmlNode> {
    let mut sections: Vec<&XmlNode> = Vec::new();
    for child in body.children() {
        if child.local() != "p" {
            continue;
        }
        if let Some(section) = child.child("pPr").and_then(|ppr| ppr.child("sectPr")) {
            sections.push(section);
        }
    }
    if let Some(section) = body.children().find(|child| child.local() == "sectPr") {
        sections.push(section);
    }
    sections
}

/// 页眉/页脚的部件路径。
///
/// **多节文档只给一套**（前端一页一页显示时页眉页脚也是一套）：从**最后一节**往前找，
/// 谁先给了 `w:type="default"` 就用谁；全都没有 default 时退回"能找到的第一个引用"。
/// 真实的企业标准就是这么分节的：封面一节、正文一节、附件一节，各自带自己的页眉页脚，
/// 若只看 `w:body/w:sectPr`（最后一节只有个页码页脚）会连"Q/…"这种标准号页眉都拿不到。
///
/// 已知取舍：**"首页不同"（`w:titlePg`）与奇偶页不区分**，同一节里优先 default。
fn story_target(body: &XmlNode, rels: &HashMap<String, RelInfo>, reference: &str) -> Option<String> {
    let sections = sections_in_order(body);
    for want_default in [true, false] {
        for section in sections.iter().rev() {
            if let Some(target) = reference_target(section, rels, reference, want_default) {
                return Some(target);
            }
        }
    }
    None
}

/// 单节里找 `w:headerReference` / `w:footerReference`：`want_default` 决定这一轮要哪种
/// `w:type`（缺省按 default 处理，宽容一点别把能读的页眉丢了）。
fn reference_target(
    section: &XmlNode,
    rels: &HashMap<String, RelInfo>,
    reference: &str,
    want_default: bool,
) -> Option<String> {
    for node in section.children_named(reference) {
        let Some(id) = node.attr_local("id") else {
            continue;
        };
        let Some(rel) = rels.get(id) else {
            continue;
        };
        let is_default = node
            .attr_local("type")
            .map(|kind| kind == "default")
            .unwrap_or(true);
        if is_default == want_default {
            return Some(rel.target.clone());
        }
    }
    None
}

/// 页眉/页脚里"什么都没有"（全是空段落）→ 当成没有。
/// 真实文档里空页脚很常见（有些只是为了让各节页脚高度对齐），前端不该为它留白。
fn story_is_empty(blocks: &[Block]) -> bool {
    blocks.iter().all(|block| match block {
        Block::Paragraph(paragraph) => paragraph.runs.iter().all(|run| run.text.trim().is_empty()),
        // 图片 / 表格 / 形状 / 文本框都算"有东西"
        _ => false,
    })
}

/// 解析页眉/页脚部件 → 内容容器（`w:hdr` / `w:ftr`）。
/// 少数生成器会多包一层，找不到就深度搜索；再找不到就把整棵树当容器。
fn parse_story_xml(xml: &str, root_local: &str) -> Option<XmlNode> {
    let root = parse_xml(xml)?;
    if root.local() == root_local {
        return Some(root);
    }
    Some(
        root.find_descendant(root_local)
            .cloned()
            .unwrap_or(root),
    )
}

/// 找页面设置（`w:sectPr`）：
/// 1. 优先文档级 `w:body/w:sectPr`（Word 把**最后一节**的设置写在这里）；
/// 2. 没有就取**最后一个**带 `w:pPr/w:sectPr` 的段落（分节符就藏在段落属性里）。
///
/// 只取一个：正文是连续流，"一页一页"显示按同一套纸张尺寸切页即可；
/// 多节文档里各节尺寸不同属于已知简化（分节处仍会画提示线）。
fn find_section_properties(body: &XmlNode) -> Option<&XmlNode> {
    if let Some(sect) = body.child("sectPr") {
        return Some(sect);
    }
    let mut found = None;
    for child in body.children() {
        if child.local() != "p" {
            continue;
        }
        if let Some(sect) = child.child("pPr").and_then(|ppr| ppr.child("sectPr")) {
            found = Some(sect);
        }
    }
    found
}

/// `w:sectPr` → 页面几何。单个字段缺失/畸形时按 A4 兜底（**不 panic**）：
/// 分页视图宁可显示一个正常的 A4，也不要因为一个坏属性把整页算成 0×0。
fn parse_page_geometry(sect: &XmlNode) -> PageGeometry {
    let size = sect.child("pgSz");
    let margins = sect.child("pgMar");

    // 纸张：必须是正数，0 / 负数 / 非数字都退回 A4
    let dimension = |value: Option<f64>, fallback: f64| match value {
        Some(value) if value > 0.0 => value,
        _ => fallback,
    };
    // 边距：允许为 0 甚至负数（Word 允许负边距），只有缺失/非数字才兜底
    let margin = |value: Option<f64>| value.unwrap_or(A4_MARGIN_PT);

    PageGeometry {
        width_pt: dimension(
            size.and_then(|node| number_f64(node.attr_local("w"))).map(twips_to_pt_f64),
            A4_WIDTH_PT,
        ),
        height_pt: dimension(
            size.and_then(|node| number_f64(node.attr_local("h"))).map(twips_to_pt_f64),
            A4_HEIGHT_PT,
        ),
        margin_top_pt: margin(
            margins
                .and_then(|node| number_f64(node.attr_local("top")))
                .map(twips_to_pt_f64),
        ),
        margin_right_pt: margin(
            margins
                .and_then(|node| number_f64(node.attr_local("right")))
                .map(twips_to_pt_f64),
        ),
        margin_bottom_pt: margin(
            margins
                .and_then(|node| number_f64(node.attr_local("bottom")))
                .map(twips_to_pt_f64),
        ),
        margin_left_pt: margin(
            margins
                .and_then(|node| number_f64(node.attr_local("left")))
                .map(twips_to_pt_f64),
        ),
        landscape: size
            .and_then(|node| node.attr_local("orient"))
            .is_some_and(|orient| orient.eq_ignore_ascii_case("landscape")),
    }
}

/// 找正文节点：正常是 `w:document/w:body`；XML 有点毛病时退化为深度搜索
fn find_body(root: &XmlNode) -> Option<&XmlNode> {
    if root.local() == "body" {
        return Some(root);
    }
    root.child("body").or_else(|| root.find_descendant("body"))
}

/// 读取部件的原始字节（图片等二进制部件用；文本部件见 [`read_part_lossy`]）
fn read_part_bytes(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Option<Vec<u8>> {
    let mut file = zip.by_name(name).ok()?;
    let mut bytes = Vec::with_capacity(file.size().min(4 * 1024 * 1024) as usize);
    file.read_to_end(&mut bytes).ok()?;
    Some(bytes)
}

/// 读取部件字节并解码成文本：UTF-8 优先，带 BOM 的 UTF-16 也认（少数生成器会写 UTF-16），
/// 最后兜底 lossy（宁可个别字符变成 U+FFFD，也不要整篇读不出来）。
fn read_part_lossy(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Option<String> {
    Some(decode_part_bytes(&read_part_bytes(zip, name)?))
}

fn decode_part_bytes(bytes: &[u8]) -> String {
    match bytes {
        [0xFF, 0xFE, rest @ ..] => decode_utf16(rest, true),
        [0xFE, 0xFF, rest @ ..] => decode_utf16(rest, false),
        _ => String::from_utf8_lossy(bytes).into_owned(),
    }
}

fn decode_utf16(bytes: &[u8], little_endian: bool) -> String {
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|pair| {
            if little_endian {
                u16::from_le_bytes([pair[0], pair[1]])
            } else {
                u16::from_be_bytes([pair[0], pair[1]])
            }
        })
        .collect();
    char::decode_utf16(units)
        .map(|item| item.unwrap_or(char::REPLACEMENT_CHARACTER))
        .collect()
}

/* ================================================================================== */
/* 八、既有能力：读字节 / 解包 / 概览 / 源码                                          */
/* ================================================================================== */

/// 读文档字节：磁盘原始字节 → 内存解密 → 明文（企业加密文件自动处理）
pub(crate) fn read_document_bytes(path: &Path) -> Result<(Vec<u8>, bool), String> {
    let (_, plain, encrypted) = office::read_raw_and_plain(path)?;
    Ok((plain, encrypted))
}

/// 把明文当 zip 打开；失败时给出能照着做的提示
fn open_zip<'a>(bytes: &'a [u8]) -> Result<ZipArchive<Cursor<&'a [u8]>>, String> {
    ZipArchive::new(Cursor::new(bytes)).map_err(|error| {
        // 走到这里说明"解密后仍不是 zip"：多半是 .doc 老格式、需要密码、或根本不是 docx
        format!(
            "这个文件不是有效的 .docx（解压失败：{error}）。\n\n\
             常见原因：\n\
             · 它是 Word 97-2003 的 .doc 老格式 —— 请用 Word / WPS 另存为 .docx\n\
             · 文件需要密码才能打开\n\
             · 文件已损坏，或不是 Word 文档（改扩展名得到的假 docx）"
        )
    })
}

/// 读取包内某个部件为字符串（缺失返回 None）
fn read_part(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Option<String> {
    let mut file = zip.by_name(name).ok()?;
    let mut text = String::new();
    file.read_to_string(&mut text).ok()?;
    Some(text)
}

/// 文档概览：部件清单 + 规模统计
#[tauri::command]
pub fn document_info(path: String) -> Result<DocumentInfo, String> {
    let (bytes, encrypted) = read_document_bytes(Path::new(&path))?;
    let mut zip = open_zip(&bytes)?;

    let mut parts = Vec::with_capacity(zip.len());
    for index in 0..zip.len() {
        let file = zip.by_index(index).map_err(|e| e.to_string())?;
        parts.push(DocumentPart {
            name: file.name().to_string(),
            size: file.size(),
        });
    }
    parts.sort_by(|a, b| b.size.cmp(&a.size));
    let images = parts
        .iter()
        .filter(|part| part.name.starts_with("word/media/"))
        .count();

    let document_xml = read_part(&mut zip, "word/document.xml").unwrap_or_default();
    let paragraphs = document_xml.matches("<w:p ").count() + document_xml.matches("<w:p>").count();
    let tables = document_xml.matches("<w:tbl>").count();

    Ok(DocumentInfo {
        encrypted,
        editable: false,
        parts,
        paragraphs,
        tables,
        images,
    })
}

/// 原始 `word/document.xml`（供「源码」视图查看，也是解析问题的排查入口）
#[tauri::command]
pub fn document_xml(path: String) -> Result<String, String> {
    let (bytes, _) = read_document_bytes(Path::new(&path))?;
    let mut zip = open_zip(&bytes)?;
    read_part(&mut zip, "word/document.xml")
        .ok_or_else(|| "包内没有 word/document.xml，可能不是 Word 文档".to_string())
}

/* ================================================================================== */
/* 九、命令：窗口化取块 / 查找                                                        */
/* ================================================================================== */

/// 取一段块（虚拟滚动窗口）。
///
/// `from` / `count` 越界按夹取处理（**不报错**）：`from > total` 返回空窗口，
/// `count` 超过 [`MAX_BLOCK_WINDOW`] 按上限截断（`total` 始终告诉前端真实块数）。
///
/// 解析 + DOM 构建是纯 CPU 活，标记 `(async)` 让 Tauri 放到线程池执行（不卡 UI 主线程），
/// Rust 签名与前端调用参数都保持不变。
#[tauri::command(async)]
pub fn document_blocks(path: String, from: usize, count: usize) -> Result<BlockPage, String> {
    let doc = load_document(&path)?;
    let total = doc.blocks.len();
    let from = from.min(total);
    let count = count.min(MAX_BLOCK_WINDOW);
    let end = from.saturating_add(count).min(total);
    Ok(BlockPage {
        total,
        from,
        blocks: doc.blocks[from..end].to_vec(),
        encrypted: doc.encrypted,
        page: doc.page.clone(),
        header: doc.header.clone(),
        footer: doc.footer.clone(),
    })
}

/// 全文查找（**块级命中**：一个块最多一条，供前端滚动定位；上限 [`MAX_FIND_HITS`] 条）。
///
/// 注意：前端 invoke 的参数名是 `matchCase`（Tauri 会把 JS 的 camelCase 参数映射到
/// Rust 的 snake_case）。
#[tauri::command(async)]
pub fn document_find(path: String, query: String, match_case: bool) -> Result<Vec<FindHit>, String> {
    let doc = load_document(&path)?;
    Ok(find_in(&doc, &query, match_case))
}

/// 取包内媒体（图片）部件，返回 **data URL**（前端可直接塞进 `<img src>`）。
///
/// - 只读且**不落临时文件**：复用 [`read_document_bytes`]（亿赛通等透明加密在内存里解密）；
/// - `media` 就是 [`ImageBlock::media`]（如 `word/media/image1.png`）；
/// - **防目录穿越**（安全边界）：只允许 `word/media/` 下的部件名，拒绝 `..`、绝对路径、盘符；
/// - 单张上限 [`MAX_MEDIA_BYTES`]，超限给中文提示，不会把几十 MB 塞进 UI。
///
/// 注意：EMF/WMF（Word 里粘贴 Excel 图表得到的矢量图）会按原 MIME 返回，但浏览器渲染不了，
/// 前端遇到这两种应显示占位卡片。
#[tauri::command(async)]
pub fn document_media(path: String, media: String) -> Result<String, String> {
    let name = validate_media_name(&media)?;
    let (bytes, _) = read_document_bytes(Path::new(&path))?;
    let mut zip = open_zip(&bytes)?;
    let data = read_media_bytes(&mut zip, &name)?;
    Ok(format!(
        "data:{};base64,{}",
        media_mime(&name),
        file::base64_encode(&data)
    ))
}

/// 校验媒体部件名。名字来自前端，这里是一道**安全边界**：
/// 我们只做 `zip.by_name` 的查表，不会碰到文件系统；但没有这层校验，前端就能拿
/// `word/document.xml`（或 `word/../word/settings.xml`）读到包内**任意部件**
/// （里面可能有批注、作者名、嵌入对象的原始数据等本不该暴露给界面的东西），
/// 所以只放行 `word/media/` 下、不含空段/`.`/`..` 的名字。
fn validate_media_name(media: &str) -> Result<String, String> {
    let name = media.trim().replace('\\', "/");
    if name.is_empty() {
        return Err("图片路径为空".to_string());
    }
    if name.starts_with('/') || name.contains(':') {
        return Err(format!("图片路径不合法（不接受绝对路径 / 盘符）：{media}"));
    }
    if name
        .split('/')
        .any(|segment| segment.is_empty() || segment == "." || segment == "..")
    {
        return Err(format!("图片路径不合法（含空段或 ..）：{media}"));
    }
    if !name.starts_with("word/media/") {
        return Err(format!(
            "只允许读取 word/media/ 下的图片，收到的是：{media}"
        ));
    }
    Ok(name)
}

/// 按扩展名给 MIME（认不出来按二进制流，前端仍可按 URL 存盘）
fn media_mime(name: &str) -> &'static str {
    match name.rsplit('.').next().unwrap_or_default().to_ascii_lowercase().as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "webp" => "image/webp",
        "svg" => "image/svg+xml",
        "tif" | "tiff" => "image/tiff",
        "ico" => "image/x-icon",
        // Word 里很常见的矢量图（粘贴 Excel 图表就是它）；浏览器渲染不了，MIME 照给
        "emf" => "image/emf",
        "wmf" => "image/wmf",
        _ => "application/octet-stream",
    }
}

/// 读媒体部件的**原始字节**（图片是二进制，不能按字符串读）。
/// zip 头里记录了原始大小，先按它拦一次，读完再兜一次（流式写入的包可能是 0）。
fn read_media_bytes(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<Vec<u8>, String> {
    let mut file = zip.by_name(name).map_err(|_| {
        format!("包里没有这张图片：{name}（可能已被删除，或文档是用「链接到文件」插入的外部图片）")
    })?;
    let declared = file.size();
    if declared > MAX_MEDIA_BYTES as u64 {
        return Err(too_large_media(name, declared as usize));
    }
    let mut data = Vec::with_capacity(declared.min(4 * 1024 * 1024) as usize);
    file.read_to_end(&mut data)
        .map_err(|error| format!("读取图片 {name} 失败：{error}"))?;
    if data.len() > MAX_MEDIA_BYTES {
        return Err(too_large_media(name, data.len()));
    }
    if data.is_empty() {
        return Err(format!("图片 {name} 是空文件（0 字节），无法显示"));
    }
    Ok(data)
}

fn too_large_media(name: &str, size: usize) -> String {
    format!(
        "图片 {name} 太大了（{}，上限 {} MB），暂不在界面里显示；用 Word / WPS 打开可以正常看到",
        human_size(size),
        MAX_MEDIA_BYTES / 1024 / 1024
    )
}

/// 字节数转成人看的大小
fn human_size(bytes: usize) -> String {
    const MB: f64 = 1024.0 * 1024.0;
    const KB: f64 = 1024.0;
    let bytes = bytes as f64;
    if bytes >= MB {
        format!("{:.1} MB", bytes / MB)
    } else if bytes >= KB {
        format!("{:.0} KB", bytes / KB)
    } else {
        format!("{bytes:.0} 字节")
    }
}

fn find_in(doc: &ParsedDocument, query: &str, match_case: bool) -> Vec<FindHit> {
    let query = query.trim();
    if query.is_empty() {
        return Vec::new();
    }
    let mut hits = Vec::new();
    for (index, text) in &doc.search_index {
        if let Some(start) = find_first(text, query, match_case) {
            hits.push(FindHit {
                block: *index,
                text: excerpt(text, start, query.chars().count()),
            });
            if hits.len() >= MAX_FIND_HITS {
                break;
            }
        }
    }
    hits
}

/// 返回**字符下标**（不是字节下标，方便截上下文）
fn find_first(text: &str, query: &str, match_case: bool) -> Option<usize> {
    if match_case {
        return text.find(query).map(|byte| text[..byte].chars().count());
    }
    // ASCII 快路径：文档里绝大多数是 ASCII 时不用按字符拆
    if text.is_ascii() && query.is_ascii() {
        let hay = text.as_bytes();
        let needle = query.as_bytes();
        if needle.len() > hay.len() {
            return None;
        }
        return (0..=hay.len() - needle.len())
            .find(|&index| hay[index..index + needle.len()].eq_ignore_ascii_case(needle));
    }
    let hay: Vec<char> = text.chars().collect();
    let needle: Vec<char> = query.chars().collect();
    if needle.is_empty() || needle.len() > hay.len() {
        return None;
    }
    (0..=hay.len() - needle.len()).find(|&index| chars_eq_ignore_case(&hay[index..index + needle.len()], &needle))
}

/// 逐字符大小写无关比较（`to_lowercase()` 对个别字符会返回多个 char，所以比较迭代器而不是字符）
fn chars_eq_ignore_case(left: &[char], right: &[char]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(a, b)| a == b || a.to_lowercase().eq(b.to_lowercase()))
}

/// 命中处的上下文片段（前后各留 [`EXCERPT_PAD`] 个字符；被截断的一侧加 `…`）
fn excerpt(text: &str, start: usize, length: usize) -> String {
    let chars: Vec<char> = text.chars().collect();
    let begin = start.saturating_sub(EXCERPT_PAD);
    let end = start.saturating_add(length).saturating_add(EXCERPT_PAD).min(chars.len());
    let mut out = String::new();
    if begin > 0 {
        out.push('…');
    }
    for ch in &chars[begin.min(chars.len())..end] {
        out.push(*ch);
    }
    if end < chars.len() {
        out.push('…');
    }
    out
}

/* ================================================================================== */
/* 测试                                                                              */
/* ================================================================================== */

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    /// 命名空间声明：真实文档里这些前缀都得有（解析按本地名匹配，写上是为了贴近真实 XML）
    const NS: &str = concat!(
        r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
        r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
        r#"xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" "#,
        r#"xmlns:v="urn:schemas-microsoft-com:vml" "#,
        r#"xmlns:o="urn:schemas-microsoft-com:office:office" "#,
        r#"xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" "#,
        r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#
    );

    /// 造一个最小可用的 docx（只有 document.xml + Content_Types），用于不依赖外部样本的单测
    fn minimal_docx(document_xml: &str) -> Vec<u8> {
        let mut buffer = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut buffer));
            let options: zip::write::FileOptions<'_, ()> =
                zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);
            writer
                .start_file("word/document.xml", options)
                .expect("写入 document.xml 失败");
            writer
                .write_all(document_xml.as_bytes())
                .expect("写入失败");
            writer.finish().expect("收尾失败");
        }
        buffer
    }

    /// 造一个多部件的 docx（内存里，不落盘），允许二进制部件（图片）
    fn docx_parts_bytes(files: Vec<(String, Vec<u8>)>) -> Vec<u8> {
        let mut buffer = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut buffer));
            let options: zip::write::FileOptions<'_, ()> =
                zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);
            for (name, body) in files {
                writer.start_file(name, options).expect("写入部件失败");
                writer.write_all(&body).expect("写入失败");
            }
            writer.finish().expect("收尾失败");
        }
        buffer
    }

    /// 造一个多部件的 docx（内存里，不落盘）
    fn docx_parts(files: Vec<(String, String)>) -> Vec<u8> {
        docx_parts_bytes(
            files
                .into_iter()
                .map(|(name, body)| (name, body.into_bytes()))
                .collect(),
        )
    }

    /// 用正文片段造 docx；`extras` 是额外的部件（styles.xml / numbering.xml / rels …）
    fn docx_with(body: &str, extras: &[(&str, &str)]) -> Vec<u8> {
        let document = format!(r#"<?xml version="1.0" encoding="UTF-8"?><w:document {NS}><w:body>{body}</w:body></w:document>"#);
        let mut files = vec![("word/document.xml".to_string(), document)];
        for (name, xml) in extras {
            files.push((name.to_string(), xml.to_string()));
        }
        docx_parts(files)
    }

    /// 直接解析字节（单测主入口）
    fn parse(bytes: &[u8]) -> ParsedDocument {
        parse_document(bytes, false).expect("解析应当成功")
    }

    /// 解析并取块
    fn blocks_of(bytes: &[u8]) -> Vec<Block> {
        parse(bytes).blocks
    }

    fn paragraph_of(block: &Block) -> &ParagraphBlock {
        match block {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("期望段落块，实际是 {other:?}"),
        }
    }

    fn table_of(block: &Block) -> &TableBlock {
        match block {
            Block::Table(table) => table,
            other => panic!("期望表格块，实际是 {other:?}"),
        }
    }

    fn label_of(block: &Block) -> &str {
        match block {
            Block::Unsupported { label, .. } => label,
            other => panic!("期望占位块，实际是 {other:?}"),
        }
    }

    /// 落盘一份测试用 docx（命令级测试必须给路径），测试结束自动删除
    struct TempDocx {
        path: std::path::PathBuf,
    }

    impl TempDocx {
        fn new(tag: &str, bytes: &[u8]) -> TempDocx {
            let dir = std::env::temp_dir().join("masteredit-docx-tests");
            std::fs::create_dir_all(&dir).expect("建临时目录失败");
            let path = dir.join(format!("{tag}-{}.docx", std::process::id()));
            std::fs::write(&path, bytes).expect("写测试文档失败");
            TempDocx { path }
        }

        fn path(&self) -> String {
            self.path.to_string_lossy().into_owned()
        }
    }

    impl Drop for TempDocx {
        fn drop(&mut self) {
            let _ = std::fs::remove_file(&self.path);
        }
    }

    /* ------------------------------ 既有测试（不要动） ------------------------------ */

    #[test]
    fn opens_minimal_docx_and_counts_blocks() {
        let xml = r#"<?xml version="1.0"?><w:document><w:body>
            <w:p><w:r><w:t>第一段</w:t></w:r></w:p>
            <w:p><w:r><w:t>第二段</w:t></w:r></w:p>
            <w:tbl><w:tr><w:tc><w:p><w:r><w:t>单元格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        </w:body></w:document>"#;
        let bytes = minimal_docx(xml);
        let mut zip = open_zip(&bytes).expect("应能作为 zip 打开");
        let text = read_part(&mut zip, "word/document.xml").expect("应能读到 document.xml");
        assert!(text.contains("第一段"), "应能读到正文内容");
        assert_eq!(text.matches("<w:p>").count(), 3, "应有 3 个段落（含表格单元格内那段）");
        assert_eq!(text.matches("<w:tbl>").count(), 1, "应有 1 个表格");
    }

    #[test]
    fn non_zip_bytes_give_actionable_error() {
        let err = open_zip(b"not a zip at all").unwrap_err();
        assert!(err.contains("不是有效的 .docx"), "错误应说明不是有效 docx：{err}");
        assert!(err.contains(".doc"), "错误应提示老格式的解决办法：{err}");
    }

    /// 用真实样本（亿赛通加密的周报）验证「解密 → 解包 → 读 XML」整条链路。
    /// 样本不存在时自动跳过，保证其它机器 / CI 上也能通过。
    #[test]
    fn decrypts_and_opens_real_encrypted_sample() {
        let sample = std::env::var("MASTEREDIT_DOCX_SAMPLE").unwrap_or_else(|_| {
            r"C:\Users\master\Documents\7.20-24周报.docx".to_string()
        });
        let path = Path::new(&sample);
        if !path.exists() {
            eprintln!("跳过：样本不存在 {sample}");
            return;
        }
        let (bytes, encrypted) = read_document_bytes(path).expect("应能读取样本");
        assert!(encrypted, "该样本应是亿赛通加密的（前 4 字节 E0 A8 91 E7）");
        assert_eq!(&bytes[0..2], b"PK", "解密后应是 zip（PK 魔数）");
        let mut zip = open_zip(&bytes).expect("解密后应能解包");
        let xml = read_part(&mut zip, "word/document.xml").expect("应能读到 document.xml");
        assert!(xml.starts_with("<?xml"), "document.xml 应是 XML 文本");
        assert!(xml.contains("<w:body>"), "应包含文档主体");
        eprintln!(
            "样本解析成功：{} 个部件，document.xml {} 字节，段落标记 {} 个，表格 {} 个",
            zip.len(),
            xml.len(),
            xml.matches("<w:p ").count() + xml.matches("<w:p>").count(),
            xml.matches("<w:tbl>").count()
        );
    }

    /* ------------------------------ 段落与 run 格式 ------------------------------ */

    #[test]
    fn parses_paragraph_runs_and_direct_formatting() {
        let body = r#"
            <w:p><w:pPr><w:jc w:val="center"/><w:ind w:left="420" w:firstLine="210"/>
              <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/></w:pPr>
              <w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:strike/>
                <w:rFonts w:ascii="Arial" w:eastAsia="宋体"/><w:sz w:val="28"/>
                <w:color w:val="ff0000"/><w:highlight w:val="yellow"/></w:rPr>
                <w:t>加粗红色</w:t></w:r>
              <w:r><w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="20"/></w:rPr><w:t>上标</w:t></w:r>
            </w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[]));
        let paragraph = paragraph_of(&blocks[0]);

        assert_eq!(paragraph.text, "加粗红色上标");
        assert_eq!(paragraph.align.as_deref(), Some("center"));
        assert_eq!(paragraph.indent_left_pt, Some(21.0), "420 twip = 21pt");
        assert_eq!(paragraph.indent_first_line_pt, Some(10.5), "210 twip = 10.5pt");
        assert_eq!(paragraph.space_before_pt, Some(6.0));
        assert_eq!(paragraph.space_after_pt, Some(12.0));
        let line = paragraph.line_spacing.as_ref().expect("应解析出行距");
        assert_eq!(line.kind, "multiple");
        assert!((line.value - 1.5).abs() < 0.001, "360/240 = 1.5 倍行距");

        let first = &paragraph.runs[0];
        assert!(first.bold && first.italic && first.underline && first.strike);
        assert_eq!(first.font.as_deref(), Some("Arial"));
        assert_eq!(first.font_east_asia.as_deref(), Some("宋体"));
        assert_eq!(first.size_pt, Some(14.0), "28 半磅 = 14pt");
        assert_eq!(first.color.as_deref(), Some("FF0000"), "颜色统一大写");
        assert_eq!(first.highlight.as_deref(), Some("FFFF00"), "具名高亮转 RRGGBB");

        let second = &paragraph.runs[1];
        assert_eq!(second.size_pt, Some(10.0));
        assert_eq!(second.vert_align.as_deref(), Some("superscript"));
    }

    #[test]
    fn resolves_doc_defaults_style_chain_and_direct_overrides() {
        // docDefaults：字号 21 半磅（10.5pt）+ 主题字体；Normal 段落样式：两端对齐
        // heading 1 基于 Normal：加粗 + 32 半磅（16pt） + 一级大纲
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:docDefaults><w:rPrDefault><w:rPr>
                <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>
                <w:sz w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>
                <w:pPr><w:jc w:val="both"/></w:pPr></w:style>
              <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
                <w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/>
                  <w:spacing w:before="240"/></w:pPr>
                <w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
              <w:style w:type="character" w:styleId="Em"><w:name w:val="Emphasis"/>
                <w:rPr><w:i/><w:color w:val="0000FF"/></w:rPr></w:style>
            </w:styles>"#
        );
        let theme = format!(
            r#"<?xml version="1.0"?><a:theme {NS}><a:themeElements><a:clrScheme name="Office"/>
              <a:fontScheme name="Office">
                <a:majorFont><a:latin typeface="等线 Light"/><a:ea typeface=""/>
                  <a:font script="Hans" typeface="等线 Light"/></a:majorFont>
                <a:minorFont><a:latin typeface="等线"/><a:ea typeface=""/>
                  <a:font script="Hans" typeface="等线"/></a:minorFont>
              </a:fontScheme></a:themeElements></a:theme>"#
        );
        let body = r#"
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
              <w:r><w:t>标题文字</w:t></w:r></w:p>
            <w:p><w:r><w:rPr><w:rStyle w:val="Em"/></w:rPr><w:t>强调</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="left"/></w:pPr>
              <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>直接格式覆盖</w:t></w:r></w:p>
            <w:p><w:pPr><w:rPr><w:rFonts w:eastAsia="仿宋"/><w:sz w:val="32"/></w:rPr></w:pPr>
              <w:r><w:t>段落标记给的字体字号</w:t></w:r></w:p>
            <w:p><w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:pPr>
              <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>run 覆盖段落标记</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(
            body,
            &[
                ("word/styles.xml", &styles),
                ("word/theme/theme1.xml", &theme),
            ],
        ));

        // 1) 标题：字号/加粗来自样式，两端对齐来自 Normal，主题字体解析成"等线"
        let heading = paragraph_of(&blocks[0]);
        assert_eq!(heading.style.as_deref(), Some("heading 1"));
        assert_eq!(heading.style_id.as_deref(), Some("Heading1"));
        assert_eq!(heading.outline_level, Some(0), "大纲级别来自样式链");
        assert_eq!(heading.align.as_deref(), Some("both"), "对齐来自 basedOn 的 Normal");
        assert_eq!(heading.space_before_pt, Some(12.0));
        let run = &heading.runs[0];
        assert_eq!(run.size_pt, Some(16.0), "样式里的 sz=32 覆盖 docDefaults 的 21");
        assert!(run.bold, "加粗来自样式");
        assert_eq!(run.font.as_deref(), Some("等线"), "主题 minorHAnsi → minorFont latin");
        assert_eq!(run.font_east_asia.as_deref(), Some("等线"), "a:ea 为空 → 用 Hans 字体");

        // 2) 字符样式：斜体 + 蓝色，字号仍走 docDefaults
        let emphasis = paragraph_of(&blocks[1]);
        let run = &emphasis.runs[0];
        assert!(run.italic, "字符样式给出斜体");
        assert_eq!(run.color.as_deref(), Some("0000FF"));
        assert_eq!(run.size_pt, Some(10.5), "docDefaults 的 21 半磅");
        assert_eq!(run.style_id.as_deref(), Some("Em"));
        assert_eq!(emphasis.align.as_deref(), Some("both"), "没写 pStyle → 用默认段落样式");

        // 3) 直接格式覆盖样式
        let overridden = paragraph_of(&blocks[2]);
        assert_eq!(overridden.align.as_deref(), Some("left"));
        assert_eq!(overridden.runs[0].size_pt, Some(12.0));
        assert!(overridden.runs[0].bold, "样式里的加粗仍然生效");

        // 4) 段落标记的 rPr（w:pPr/w:rPr）是段落文字的默认格式：run 没写就继承它
        let from_mark = paragraph_of(&blocks[3]);
        assert_eq!(from_mark.runs[0].size_pt, Some(16.0), "32 半磅来自段落标记的 rPr");
        assert_eq!(from_mark.runs[0].font_east_asia.as_deref(), Some("仿宋"));

        // 5) run 自己的 rPr 仍然盖过段落标记
        let run_wins = paragraph_of(&blocks[4]);
        assert_eq!(run_wins.runs[0].size_pt, Some(9.0), "18 半磅来自 run");
    }

    #[test]
    fn cyclic_style_inheritance_does_not_hang() {
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:basedOn w:val="B"/>
                <w:rPr><w:sz w:val="20"/></w:rPr></w:style>
              <w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/><w:basedOn w:val="A"/>
                <w:rPr><w:b/></w:rPr></w:style>
            </w:styles>"#
        );
        let body = r#"<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>成环</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[("word/styles.xml", &styles)]));
        let paragraph = paragraph_of(&blocks[0]);
        assert_eq!(paragraph.runs[0].size_pt, Some(10.0), "本样式的 sz 生效");
        assert!(paragraph.runs[0].bold, "父样式的 b 也能拿到（成环被截断但没崩）");
    }

    #[test]
    fn decodes_entities_tabs_breaks_and_symbols() {
        let body = r#"
            <w:p><w:r><w:t xml:space="preserve">a &amp; b &lt;tag&gt; &#x4E2D; </w:t></w:r>
              <w:r><w:t>换</w:t><w:br/><w:t>行</w:t></w:r>
              <w:r><w:tab/><w:t>制表</w:t></w:r>
              <w:r><w:sym w:font="Wingdings" w:char="F0E0"/></w:r>
            </w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[]));
        let paragraph = paragraph_of(&blocks[0]);
        assert!(paragraph.text.starts_with("a & b <tag> 中 "), "实体应还原：{:?}", paragraph.text);
        assert!(paragraph.text.contains("换\n行"), "软换行应是 \\n：{:?}", paragraph.text);
        assert!(paragraph.text.contains("\t制表"), "制表符应保留：{:?}", paragraph.text);
        assert!(paragraph.text.contains('\u{F0E0}'), "w:sym 应按码位还原");
    }

    /* ------------------------------ 编号与多级列表 ------------------------------ */

    #[test]
    fn parses_multilevel_numbering_and_bullets() {
        let numbering = format!(
            r#"<?xml version="1.0"?><w:numbering {NS}>
              <w:abstractNum w:abstractNumId="7">
                <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
                  <w:suff w:val="space"/><w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl>
                <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/>
                  <w:pPr><w:ind w:left="840" w:hanging="420"/></w:pPr></w:lvl>
                <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%3、"/></w:lvl>
              </w:abstractNum>
              <w:abstractNum w:abstractNumId="9">
                <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
              </w:abstractNum>
              <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
              <w:num w:numId="4"><w:abstractNumId w:val="9"/></w:num>
            </w:numbering>"#
        );
        let item = |num: u32, level: u32, text: &str| {
            format!(
                r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num}"/></w:numPr></w:pPr>
                   <w:r><w:t>{text}</w:t></w:r></w:p>"#
            )
        };
        let body = format!(
            "{}{}{}{}{}{}",
            item(3, 0, "一"),
            item(3, 1, "一甲"),
            item(3, 1, "一乙"),
            item(3, 2, "深"),
            item(3, 0, "二"),
            item(4, 0, "符号项"),
        );
        let blocks = blocks_of(&docx_with(&body, &[("word/numbering.xml", &numbering)]));

        let prefixes: Vec<String> = blocks
            .iter()
            .map(|block| {
                paragraph_of(block)
                    .list
                    .as_ref()
                    .map(|list| list.prefix.clone())
                    .unwrap_or_else(|| "<无编号>".to_string())
            })
            .collect();
        assert_eq!(
            prefixes,
            vec!["1.", "a)", "b)", "一、", "2.", "•"],
            "多级编号要算对父级：第二层用字母、第三层用中文数字、回到一级后继续 2."
        );

        let first = paragraph_of(&blocks[0]).list.clone().expect("应有列表信息");
        assert!(first.ordered);
        assert_eq!(first.level, 0);
        assert_eq!(first.suffix, "space");
        assert_eq!(first.format, "decimal");
        // 编号级别自带的缩进参与段落层叠（420 twip = 21pt，悬挂缩进 → 首行 -21pt）
        let paragraph = paragraph_of(&blocks[0]);
        assert_eq!(paragraph.indent_left_pt, Some(21.0));
        assert_eq!(paragraph.indent_first_line_pt, Some(-21.0));

        let bullet = paragraph_of(&blocks[5]).list.clone().expect("项目符号也是列表");
        assert!(!bullet.ordered, "bullet 应标成无序");
        assert_eq!(bullet.prefix, "•");
    }

    #[test]
    fn numbering_restarts_per_num_id_and_supports_overrides() {
        // 同一个 abstractNum 被两个 numId 复用：两条列表各自从 1 开始数
        let numbering = format!(
            r#"<?xml version="1.0"?><w:numbering {NS}>
              <w:abstractNum w:abstractNumId="1">
                <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1、"/></w:lvl>
              </w:abstractNum>
              <w:num w:numId="10"><w:abstractNumId w:val="1"/></w:num>
              <w:num w:numId="11"><w:abstractNumId w:val="1"/>
                <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>
            </w:numbering>"#
        );
        let item = |num: u32, text: &str| {
            format!(
                r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr></w:pPr>
                   <w:r><w:t>{text}</w:t></w:r></w:p>"#
            )
        };
        let body = format!(
            "{}{}{}{}",
            item(10, "A"),
            item(10, "B"),
            item(11, "C"),
            item(11, "D")
        );
        let blocks = blocks_of(&docx_with(&body, &[("word/numbering.xml", &numbering)]));
        let prefixes: Vec<String> = blocks
            .iter()
            .map(|block| paragraph_of(block).list.clone().unwrap().prefix)
            .collect();
        assert_eq!(
            prefixes,
            vec!["I、", "II、", "V、", "VI、"],
            "numId=10 从 I 开始；numId=11 用 startOverride=5 从 V 开始"
        );
    }

    /* ------------------------------ 表格 ------------------------------ */

    #[test]
    fn parses_table_merges_widths_borders_and_nested_table() {
        let body = r#"
            <w:tbl>
              <w:tblPr><w:tblStyle w:val="a1"/><w:tblW w:w="6000" w:type="dxa"/><w:jc w:val="center"/>
                <w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
                  <w:bottom w:val="nil"/><w:right w:val="none"/></w:tblBorders>
              </w:tblPr>
              <w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>
              <w:tr><w:trPr><w:trHeight w:val="400"/><w:tblHeader/></w:trPr>
                <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:gridSpan w:val="2"/>
                    <w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr>
                  <w:p><w:r><w:t>合并两列</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr>
                  <w:p><w:r><w:t>第三列</w:t></w:r></w:p>
                  <w:p><w:r><w:t>第二段</w:t></w:r></w:p>
                  <w:tbl><w:tblGrid><w:gridCol w:w="600"/></w:tblGrid>
                    <w:tr><w:tc><w:p><w:r><w:t>嵌套</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
                </w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
                <w:tc><w:tcPr><w:tcBorders><w:top w:val="nil"/></w:tcBorders></w:tcPr>
                  <w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr/></w:tc>
              </w:tr>
            </w:tbl>"#;
        let blocks = blocks_of(&docx_with(body, &[]));
        assert_eq!(blocks.len(), 1, "一个表格就是一个顶层块");
        let table = table_of(&blocks[0]);

        assert_eq!(table.rows.len(), 3);
        assert_eq!(table.align.as_deref(), Some("center"));
        assert_eq!(table.style_id.as_deref(), Some("a1"));
        assert!(table.borders, "单元格上的可见边框也算表格有边框");
        assert_eq!(
            table.columns,
            vec![100.0, 100.0, 200.0],
            "1500 twip = 100px，3000 twip = 200px"
        );
        assert_eq!(table.width_px, Some(400.0), "6000 twip = 400px");

        // 第一行：横跨两列 + 底纹 + 行高 + 表头
        let row = &table.rows[0];
        assert!(row.header, "w:tblHeader 应标成表头行");
        assert!((row.height_px.unwrap() - 26.6667).abs() < 0.01, "400 twip ≈ 26.67px");
        assert_eq!(row.cells.len(), 2);
        assert_eq!(row.cells[0].grid_span, 2);
        assert_eq!(row.cells[0].text, "合并两列");
        assert_eq!(row.cells[0].shading.as_deref(), Some("DDEEFF"));
        assert_eq!(row.cells[0].width_px, Some(200.0));
        assert_eq!(row.cells[1].v_align.as_deref(), Some("center"));
        assert_eq!(
            row.cells[1].text, "第三列\n第二段\n嵌套",
            "单元格内多段落/嵌套表格的文本用换行连接"
        );
        assert_eq!(row.cells[1].blocks.len(), 3, "两个段落 + 一个嵌套表格");
        let nested = table_of(&row.cells[1].blocks[2]);
        assert_eq!(nested.rows[0].cells[0].text, "嵌套");

        // 单元格边框：自己没写就用表格级兜底（top=single 可见，bottom=nil 不可见）
        let inherited = row.cells[1].borders;
        assert!(inherited.top && inherited.left && !inherited.bottom && !inherited.right);

        // 纵向合并
        assert_eq!(table.rows[1].cells[0].v_merge, VMerge::Restart);
        assert_eq!(table.rows[2].cells[0].v_merge, VMerge::Continue);
        assert_eq!(table.rows[1].cells[2].v_merge, VMerge::Continue);
        assert_eq!(table.rows[0].cells[0].v_merge, VMerge::None);

        // 单元格自己写的 nil 覆盖表格级边框
        let explicit = table.rows[2].cells[1].borders;
        assert!(!explicit.top, "自己写了 nil 就按 nil");
        assert!(explicit.left, "没写的边继续沿用表格级");

        // 空单元格也要保留（否则前端表格会错列）
        assert!(table.rows[2].cells[2].blocks.is_empty());
        assert_eq!(table.rows[2].cells[2].text, "");
    }

    /* ------------------------------ 图片与占位块 ------------------------------ */

    #[test]
    fn resolves_images_from_relationships() {
        let rels = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
              <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../word/media/image2.jpeg"/>
              <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
            </Relationships>"#
        );
        let body = r#"
            <w:p><w:r><w:rPr><w:noProof/></w:rPr><w:drawing>
              <wp:inline><wp:extent cx="914400" cy="457200"/>
                <wp:docPr id="1" name="图片 1" descr="示意图"/>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic>
                </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
            <w:p><w:r><w:t>带图段落</w:t></w:r>
              <w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic><pic:blipFill><a:blip r:embed="rId6"/></pic:blipFill></pic:pic>
                </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
            <w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic><pic:blipFill><a:blip r:embed="rId404"/></pic:blipFill></pic:pic>
              </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(
            body,
            &[("word/_rels/document.xml.rels", &rels)],
        ));

        let image = match &blocks[0] {
            Block::Image(image) => image,
            other => panic!("整段只有图片时应产出图片块，实际 {other:?}"),
        };
        assert_eq!(image.media, "word/media/image1.png");
        assert_eq!(image.name.as_deref(), Some("图片 1"));
        assert_eq!(image.alt.as_deref(), Some("示意图"));
        assert_eq!((image.width_px, image.height_px), (96.0, 48.0), "914400 EMU = 96px");

        // 段落里有文字 + 图片：文字成段落块，图片紧跟其后
        let paragraph = paragraph_of(&blocks[1]);
        assert_eq!(paragraph.text, "带图段落");
        let second = match &blocks[2] {
            Block::Image(image) => image,
            other => panic!("期望图片块，实际 {other:?}"),
        };
        assert_eq!(second.media, "word/media/image2.jpeg", "../word/media → 归一化");
        assert_eq!((second.width_px, second.height_px), (1.0, 1.0));

        // 关系缺失：给占位说明而不是静默丢
        let missing = label_of(&blocks[3]);
        assert!(missing.contains("图片"), "关系缺失的图片要有明确说明：{missing}");
        let detail = match &blocks[3] {
            Block::Unsupported { detail, .. } => detail,
            _ => unreachable!(),
        };
        assert!(detail.contains("找不到"), "说明里要讲清原因：{detail}");
    }

    #[test]
    fn emits_placeholders_for_unsupported_objects() {
        let rels = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/ole.png"/>
              <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/altChunk" Target="afchunk.html"/>
            </Relationships>"#
        );
        let body = r#"
            <w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
            <w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <a:blip r:embed="rId99"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
            <w:p><w:r><w:object>
              <v:shape id="ole预览" style="width:72pt;height:36pt"><v:imagedata r:id="rId6"/></v:shape>
              <o:OLEObject ProgID="Equation.DSMT4"/></w:object></w:r></w:p>
            <w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>
              <w:p><w:r><w:t>框内文字</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>
            <w:p><m:oMathPara><m:oMath><m:r><m:t>E</m:t></m:r><m:r><m:t>=mc</m:t></m:r></m:oMath></m:oMathPara></w:p>
            <w:p><w:r><w:t>普通段落</w:t></w:r></w:p>
            <w:altChunk r:id="rId9"/>"#;
        let blocks = blocks_of(&docx_with(
            body,
            &[("word/_rels/document.xml.rels", &rels)],
        ));

        let shape: Vec<&str> = blocks
            .iter()
            .map(|block| match block {
                Block::Paragraph(_) => "段落",
                Block::Table(_) => "表格",
                Block::Image(_) => "图片",
                Block::Shape(_) => "形状",
                Block::TextBox(_) => "文本框",
                Block::PageBreak => "分页",
                Block::Unsupported { .. } => "占位",
            })
            .collect();
        assert_eq!(
            shape,
            vec!["占位", "占位", "图片", "占位", "文本框", "占位", "段落", "占位"],
            "每个不认识的节点都要有落点，绝不能静默消失（文本框现在有真内容）"
        );

        assert!(label_of(&blocks[0]).contains("SmartArt"), "SmartArt 要有中文说明");
        assert!(label_of(&blocks[1]).contains("图表"), "图表要有中文说明");
        // OLE：先给预览图（文档里内嵌的），再给占位说明
        let preview_image = match &blocks[2] {
            Block::Image(image) => image,
            other => panic!("OLE 对象前应有预览图块，实际 {other:?}"),
        };
        assert_eq!(preview_image.media, "word/media/ole.png");
        assert_eq!(
            (preview_image.width_px, preview_image.height_px),
            (96.0, 48.0),
            "72pt × 36pt → 96px × 48px"
        );
        assert!(label_of(&blocks[3]).contains("OLE"), "OLE 对象给占位");
        let ole_detail = match &blocks[3] {
            Block::Unsupported { detail, .. } => detail,
            _ => unreachable!(),
        };
        assert!(ole_detail.contains("预览图"), "OLE 应说明显示的是预览图：{ole_detail}");
        assert!(ole_detail.contains("Equation.DSMT4"), "OLE 应报出 ProgID：{ole_detail}");
        // 文本框现在是**真内容**（这块行为已由 renders_text_boxes_with_real_content 详测）
        let text_box = match &blocks[4] {
            Block::TextBox(text_box) => text_box,
            other => panic!("VML 文本框应产出文本框块：{other:?}"),
        };
        assert_eq!(paragraph_of(&text_box.blocks[0]).text, "框内文字");
        assert!(label_of(&blocks[5]).contains("公式"));
        let math_detail = match &blocks[5] {
            Block::Unsupported { detail, .. } => detail,
            _ => unreachable!(),
        };
        assert!(math_detail.contains("E=mc"), "公式占位里回显纯文本：{math_detail}");
        assert!(label_of(&blocks[7]).contains("altChunk"));
    }

    #[test]
    fn handles_page_break_image_paragraphs_and_empty_paragraphs() {
        let rels = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/b.png"/>
            </Relationships>"#
        );
        let drawing = |rid: &str| {
            format!(
                r#"<w:r><w:drawing><wp:inline><wp:extent cx="19050" cy="9525"/>
                  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic><pic:blipFill><a:blip r:embed="{rid}"/></pic:blipFill></pic:pic>
                  </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
            )
        };
        let body = format!(
            r#"{empty}{page}{two_images}{editable}{before}{mixed}{section}"#,
            empty = r#"<w:p/>"#,
            page = r#"<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>"#,
            two_images = format!("<w:p>{}{}</w:p>", drawing("rId1"), drawing("rId2")),
            editable = r#"<w:p><w:pPr><w:pStyle w:val="a9"/></w:pPr><w:r><w:t>文字</w:t></w:r></w:p>"#,
            before = r#"<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>段前分页</w:t></w:r></w:p>"#,
            mixed = r#"<w:p><w:r><w:t>前</w:t><w:br w:type="page"/><w:t>后</w:t></w:r></w:p>"#,
            section = r#"<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>分节</w:t></w:r></w:p>"#,
        );
        let blocks = blocks_of(&docx_with(&body, &[("word/_rels/document.xml.rels", &rels)]));

        // 空段落也要成块（前端靠它撑段间距）
        let empty = paragraph_of(&blocks[0]);
        assert_eq!(empty.text, "");
        assert!(empty.runs.is_empty());
        // 只含分页符的段落 → 分页块
        assert!(matches!(blocks[1], Block::PageBreak), "实际 {:?}", blocks[1]);
        // 只有图片的段落 → 一张图一个块
        assert!(matches!(blocks[2], Block::Image(_)));
        assert!(matches!(blocks[3], Block::Image(_)));
        assert_eq!(
            match &blocks[2] {
                Block::Image(image) => (image.width_px, image.height_px),
                _ => unreachable!(),
            },
            (2.0, 1.0),
            "19050 EMU = 2px"
        );
        // 普通段落
        assert_eq!(paragraph_of(&blocks[4]).text, "文字");
        // 段前分页标记
        assert!(paragraph_of(&blocks[5]).page_break_before);
        // 段落内混排分页符：保留成一段，但标记出来（内容不丢）
        let mixed = paragraph_of(&blocks[6]);
        assert!(mixed.page_break, "混排分页符要标记");
        assert_eq!(mixed.text, "前\n后");
        // 分节符
        let section = paragraph_of(&blocks[7]);
        assert_eq!(section.section_break.as_deref(), Some("continuous"));
    }

    /* ------------------------------ 损坏输入不崩 ------------------------------ */

    #[test]
    fn broken_xml_never_panics() {
        // 1) 结束标签不匹配 + 裸 & + 未闭合标签
        let body = r#"<w:p><w:r><w:t>甲 & 乙</w:t></w:r></w:p><w:p><w:r><w:t>未闭合"#;
        let blocks = blocks_of(&docx_with(body, &[]));
        assert!(!blocks.is_empty(), "损坏的 XML 也要尽量读出内容");
        let text: String = blocks
            .iter()
            .map(block_plain_text)
            .collect::<Vec<_>>()
            .join("");
        assert!(text.contains("甲 & 乙"), "裸 & 不该整篇崩：{text:?}");

        // 2) 完全没有正文节点
        let err = parse_document(&docx_with("", &[]).len().to_string().into_bytes(), false);
        assert!(err.is_err(), "非 zip 字节应返回 Err 而不是 panic");

        // 3) 有 body 但空 → 空块列表
        let empty = blocks_of(&docx_with("", &[]));
        assert!(empty.is_empty());

        // 4) styles.xml / numbering.xml 是垃圾 → 不影响正文
        let blocks = blocks_of(&docx_with(
            r#"<w:p><w:pPr><w:pStyle w:val="没有这个样式"/><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr>
               <w:r><w:t>正文</w:t></w:r></w:p>"#,
            &[
                ("word/styles.xml", "<<<这不是 XML"),
                ("word/numbering.xml", "<w:numbering><w:abstractNum 半截"),
            ],
        ));
        assert_eq!(paragraph_of(&blocks[0]).text, "正文");
        // 编号级别查不到时用兜底定义（"1."），而不是让前缀变空
        let list = paragraph_of(&blocks[0]).list.clone().expect("仍应有列表信息");
        assert_eq!(list.prefix, "1.");

        // 5) 属性里带非法实体
        let blocks = blocks_of(&docx_with(
            r#"<w:p><w:pPr><w:jc w:val="a&b"/></w:pPr><w:r><w:t>属性异常</w:t></w:r></w:p>"#,
            &[],
        ));
        assert_eq!(paragraph_of(&blocks[0]).text, "属性异常");
    }

    #[test]
    fn table_and_cell_counts_are_consistent_for_odd_input() {
        // 空表格、没有 tblGrid、单元格里什么都没有
        let body = r#"<w:tbl/><w:tbl><w:tr/></w:tbl><w:tbl><w:tr><w:tc/></w:tr></w:tbl>"#;
        let blocks = blocks_of(&docx_with(body, &[]));
        assert_eq!(blocks.len(), 3);
        for block in &blocks {
            let table = table_of(block);
            assert!(table.columns.is_empty());
        }
        assert!(table_of(&blocks[1]).rows[0].cells.is_empty());
        assert_eq!(table_of(&blocks[2]).rows[0].cells.len(), 1);
        assert_eq!(table_of(&blocks[2]).rows[0].cells[0].grid_span, 1);
    }

    #[test]
    fn decodes_utf16_and_gbk_like_bytes_without_panicking() {
        // UTF-16LE + BOM
        let mut utf16 = vec![0xFF, 0xFE];
        for unit in "中文".encode_utf16() {
            utf16.extend_from_slice(&unit.to_le_bytes());
        }
        assert_eq!(decode_part_bytes(&utf16), "中文");

        // UTF-16BE + BOM
        let mut utf16 = vec![0xFE, 0xFF];
        for unit in "AB".encode_utf16() {
            utf16.extend_from_slice(&unit.to_be_bytes());
        }
        assert_eq!(decode_part_bytes(&utf16), "AB");

        // 非法 UTF-8 也不 panic（UTF-16 里的孤立代理项 → 替换字符）
        assert_eq!(decode_part_bytes(&[0xFF, 0xFE, 0x00, 0xD8]), "\u{FFFD}");
        assert_eq!(decode_part_bytes(&[0x41, 0xFF, 0x42]), "A\u{FFFD}B");
        assert_eq!(decode_part_bytes(b"<w:p/>"), "<w:p/>");
    }

    /* ------------------------------ 命令：窗口化与查找 ------------------------------ */

    #[test]
    fn document_blocks_windows_and_clamps_out_of_range() {
        let mut body = String::new();
        for index in 0..12 {
            body.push_str(&format!("<w:p><w:r><w:t>第{index}段</w:t></w:r></w:p>"));
        }
        let doc = TempDocx::new("window", &docx_with(&body, &[]));
        let path = doc.path();

        let page = document_blocks(path.clone(), 3, 4).expect("取块应成功");
        assert_eq!(page.total, 12);
        assert_eq!(page.from, 3);
        assert_eq!(page.blocks.len(), 4);
        assert_eq!(paragraph_of(&page.blocks[0]).text, "第3段");
        assert!(!page.encrypted);

        // 越界：from 超出总数 → 空窗口而不是报错
        let page = document_blocks(path.clone(), 999, 10).expect("越界不应报错");
        assert_eq!(page.from, 12);
        assert!(page.blocks.is_empty());

        // count 超上限 → 夹到 MAX_BLOCK_WINDOW（这里文档更小，所以就是全部）
        let page = document_blocks(path.clone(), 0, usize::MAX).expect("超大 count 不应报错");
        assert_eq!(page.blocks.len(), 12);

        // count = 0 → 空窗口
        let page = document_blocks(path.clone(), 0, 0).expect("count=0 不应报错");
        assert!(page.blocks.is_empty());
        assert_eq!(page.total, 12);

        // 文件不存在 → Err（不是 panic）
        let missing = document_blocks(doc.path().replace(".docx", "-不存在.docx"), 0, 10);
        assert!(missing.is_err());
    }

    #[test]
    fn document_find_is_block_level_case_insensitive_and_capped() {
        let body = r#"
            <w:p><w:r><w:t>Hello World</w:t></w:r></w:p>
            <w:p><w:r><w:t>第二段里有 hello 两次：hello</w:t></w:r></w:p>
            <w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格 hello</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
        let doc = TempDocx::new("find", &docx_with(body, &[]));
        let path = doc.path();

        let hits = document_find(path.clone(), "hello".to_string(), false).expect("查找应成功");
        assert_eq!(hits.len(), 3, "每个块最多一条命中（第二段里的两次 hello 只算一条）");
        assert_eq!(
            hits.iter().map(|hit| hit.block).collect::<Vec<_>>(),
            vec![0, 1, 2]
        );
        assert!(hits[1].text.contains("hello"));
        assert!(
            hits[1].text.matches("hello").count() >= 1,
            "片段要包含命中处：{}",
            hits[1].text
        );

        let hits = document_find(path.clone(), "hello".to_string(), true).expect("查找应成功");
        assert_eq!(
            hits.iter().map(|hit| hit.block).collect::<Vec<_>>(),
            vec![1, 2],
            "区分大小写时只命中「hello」小写的两处（第一段是 Hello）"
        );

        // 表格是整表一块：命中定位到表格那一块
        let hits = document_find(path.clone(), "表格 hello".to_string(), false).unwrap();
        assert_eq!(hits.len(), 1);
        assert_eq!(hits[0].block, 2);

        // 空查询 / 查不到
        assert!(document_find(path.clone(), "   ".to_string(), false).unwrap().is_empty());
        assert!(document_find(path, "不存在的词".to_string(), false).unwrap().is_empty());
    }

    #[test]
    fn document_find_caps_hits_at_five_hundred() {
        let mut body = String::new();
        for index in 0..600 {
            body.push_str(&format!("<w:p><w:r><w:t>命中 {index}</w:t></w:r></w:p>"));
        }
        let doc = TempDocx::new("findcap", &docx_with(&body, &[]));
        let hits = document_find(doc.path(), "命中".to_string(), false).unwrap();
        assert_eq!(hits.len(), MAX_FIND_HITS, "命中上限 {MAX_FIND_HITS} 条");
        assert_eq!(hits[0].block, 0);
        assert_eq!(hits[MAX_FIND_HITS - 1].block, MAX_FIND_HITS - 1);
    }

    #[test]
    fn excerpt_keeps_context_around_match() {
        let text = "abcdefghij".repeat(20);
        let snippet = excerpt(&text, 100, 3);
        assert!(snippet.starts_with('…') && snippet.ends_with('…'));
        assert!(snippet.contains("abc"), "片段里应包含命中处内容");
        let short = excerpt("短文本", 0, 1);
        assert_eq!(short, "短文本", "短文本不加省略号");
    }

    /* ------------------------------ 给前端的 JSON 形状（契约冻结） ------------------------------ */

    /// 前端会照着这些字段名写 TS 类型，所以把它们钉死在测试里：
    /// 块用 `kind` 判别（`paragraph` / `table` / `image` / `pageBreak` / `unsupported`），
    /// 字段名一律 camelCase。
    #[test]
    fn serializes_blocks_with_frontend_friendly_field_names() {
        let body = r#"
            <w:p><w:pPr><w:ind w:left="420"/></w:pPr>
              <w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="24"/><w:b/></w:rPr>
                <w:t>文字</w:t></w:r></w:p>
            <w:p><w:r><w:br w:type="page"/></w:r></w:p>
            <w:tbl><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
              <w:tr><w:tc><w:tcPr><w:gridSpan w:val="1"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>"#;
        let parsed = parse(&docx_with(body, &[]));
        let page = BlockPage {
            total: parsed.blocks.len(),
            from: 0,
            blocks: parsed.blocks.clone(),
            encrypted: true,
            page: parsed.page.clone(),
            header: parsed.header.clone(),
            footer: parsed.footer.clone(),
        };
        let json = serde_json::to_value(&page).expect("应能序列化成 JSON");

        assert_eq!(json["total"], 3);
        assert!(
            json["page"].is_null(),
            "这份样本没有 w:sectPr，page 应是 null（前端按 A4 兜底）"
        );
        assert_eq!(json["from"], 0);
        assert_eq!(json["encrypted"], true);

        assert_eq!(json["blocks"][0]["kind"], "paragraph");
        assert_eq!(json["blocks"][0]["text"], "文字");
        assert_eq!(json["blocks"][0]["indentLeftPt"], 21.0);
        assert_eq!(json["blocks"][0]["runs"][0]["bold"], true);
        assert_eq!(json["blocks"][0]["runs"][0]["fontEastAsia"], "宋体");
        assert_eq!(json["blocks"][0]["runs"][0]["sizePt"], 12.0);
        assert_eq!(json["blocks"][1]["kind"], "pageBreak");
        assert_eq!(json["blocks"][2]["kind"], "table");
        assert_eq!(json["blocks"][2]["columns"][0], 100.0);
        assert_eq!(json["blocks"][2]["rows"][0]["cells"][0]["gridSpan"], 1);
        assert_eq!(json["blocks"][2]["rows"][0]["cells"][0]["vMerge"], "none");
        assert!(json["blocks"][2]["rows"][0]["cells"][0]["borders"]["top"].is_boolean());

        // 占位块：label + detail 都要在
        let json = serde_json::to_value(unsupported("测试占位", "中文说明")).expect("应能序列化");
        assert_eq!(json["kind"], "unsupported");
        assert_eq!(json["label"], "测试占位");
        assert_eq!(json["detail"], "中文说明");

        // 图片块：媒体路径 + px 尺寸
        let json = serde_json::to_value(Block::Image(ImageBlock {
            media: "word/media/a.png".to_string(),
            name: Some("图片 1".to_string()),
            alt: None,
            width_px: 10.0,
            height_px: 5.0,
        }))
        .expect("应能序列化");
        assert_eq!(json["kind"], "image");
        assert_eq!(json["media"], "word/media/a.png");
        assert_eq!(json["widthPx"], 10.0);
        assert_eq!(json["heightPx"], 5.0);
        assert!(json["alt"].is_null());

        // 列表信息：前缀文本 + 有序标记
        let numbered = parse(&docx_with(
            r#"<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>项</w:t></w:r></w:p>"#,
            &[(
                "word/numbering.xml",
                r#"<?xml version="1.0"?><w:numbering>
                     <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
                       <w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
                     <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#,
            )],
        ));
        let json = serde_json::to_value(&numbered.blocks[0]).expect("应能序列化");
        assert_eq!(json["kind"], "paragraph");
        assert_eq!(json["list"]["prefix"], "1.");
        assert_eq!(json["list"]["ordered"], true);
        assert_eq!(json["list"]["numId"], 1);
        assert_eq!(json["list"]["level"], 0);

        // 查找命中的形状
        let json = serde_json::to_value(FindHit {
            block: 7,
            text: "…上下文…".to_string(),
        })
        .expect("应能序列化");
        assert_eq!(json["block"], 7);
        assert_eq!(json["text"], "…上下文…");
    }

    /* ------------------------------ 图片：document_media（data URL） ------------------------------ */

    /// 1×1 透明 PNG（67 字节的最小合法 PNG）——base64 见 [`TINY_PNG_BASE64`]
    const TINY_PNG: [u8; 67] = [
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82,
    ];
    /// 上面那张 PNG 的 base64（用 PowerShell 的 `[Convert]::ToBase64String` 独立算出来的，
    /// 所以这条断言是"跟外部实现对答案"，而不是自己跟自己对）
    const TINY_PNG_BASE64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
    /// PNG 魔数（89 50 4E 47 0D 0A 1A 0A）
    const PNG_MAGIC: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];

    /// 一个只有段落的正文 + 一张图片部件的最小 docx
    fn docx_with_png() -> Vec<u8> {
        docx_parts_bytes(vec![
            (
                "word/document.xml".to_string(),
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8"?><w:document {NS}><w:body>{}</w:body></w:document>"#,
                    r#"<w:p><w:r><w:t>带图</w:t></w:r></w:p>"#
                )
                .into_bytes(),
            ),
            ("word/media/pixel.png".to_string(), TINY_PNG.to_vec()),
        ])
    }

    #[test]
    fn document_media_returns_data_url_with_exact_bytes() {
        let doc = TempDocx::new("media-png", &docx_with_png());
        let url = document_media(doc.path(), "word/media/pixel.png".to_string())
            .expect("应能取到图片");

        let payload = url
            .strip_prefix("data:image/png;base64,")
            .unwrap_or_else(|| panic!("应是 PNG 的 data URL：{}", &url[..40.min(url.len())]));
        assert_eq!(payload, TINY_PNG_BASE64, "编码结果应与外部独立算出的 base64 一致");

        // 解码回来，确认字节一模一样（前端拿到的是原图）
        let decoded = file::base64_decode(payload).expect("应能解码 data URL");
        assert_eq!(decoded, TINY_PNG.to_vec());
        assert_eq!(decoded[0..8], PNG_MAGIC, "PNG 魔数");
    }

    #[test]
    fn document_media_rejects_traversal_and_paths_outside_media() {
        let doc = TempDocx::new(
            "media-bad",
            &docx_with(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#, &[]),
        );
        for (bad, expect) in [
            ("../word/document.xml", ".."),
            ("word/media/../../word/document.xml", ".."),
            ("..\\word\\document.xml", ".."),
            ("/etc/passwd", "绝对路径"),
            ("C:\\Windows\\win.ini", "绝对路径"),
            ("//server/share/a.png", "绝对路径"),
            ("word/document.xml", "只允许"),
            ("word/media/", "空段"),
            ("word/media/./a.png", "空段"),
            ("  ", "为空"),
        ] {
            let result = document_media(doc.path(), bad.to_string());
            let error = match result {
                Ok(url) => panic!("{bad:?} 必须被拒绝，却返回了 {}", &url[..40.min(url.len())]),
                Err(error) => error,
            };
            assert!(
                error.contains(expect) || error.contains("不合法") || error.contains("只允许"),
                "{bad:?} 的报错要说清原因（期望含 {expect:?}）：{error}"
            );
        }

        // 合法名字能通过校验（这里包里没有这个部件，所以错误是"包里没有"）
        let error = document_media(doc.path(), "word/media/pixel.png".to_string()).unwrap_err();
        assert!(error.contains("包里没有"), "不存在的图片要有可读报错：{error}");
        assert!(
            error.contains("链接到文件"),
            "要提示可能是外部链接图片：{error}"
        );
    }

    #[test]
    fn document_media_reports_broken_input_cleanly() {
        // 1) 空图片部件
        let files = docx_parts_bytes(vec![
            (
                "word/document.xml".to_string(),
                format!(r#"<?xml version="1.0"?><w:document {NS}><w:body/></w:document>"#).into_bytes(),
            ),
            ("word/media/empty.png".to_string(), Vec::new()),
        ]);
        let doc = TempDocx::new("media-empty", &files);
        let error = document_media(doc.path(), "word/media/empty.png".to_string()).unwrap_err();
        assert!(error.contains("空文件"), "{error}");

        // 2) 不是 docx（也不是 zip）
        let dir = std::env::temp_dir().join("masteredit-docx-tests");
        std::fs::create_dir_all(&dir).expect("建临时目录失败");
        let bad = dir.join(format!("media-notzip-{}.docx", std::process::id()));
        std::fs::write(&bad, b"not a zip").expect("写文件失败");
        let error =
            document_media(bad.to_string_lossy().into_owned(), "word/media/a.png".to_string())
                .unwrap_err();
        let _ = std::fs::remove_file(&bad);
        assert!(error.contains("不是有效的 .docx"), "{error}");

        // 3) 文件不存在
        let missing = dir.join("media-不存在.docx");
        let error =
            document_media(missing.to_string_lossy().into_owned(), "word/media/a.png".to_string())
                .unwrap_err();
        assert!(!error.is_empty(), "应给出可读错误：{error}");

        // 4) 超限保护（不真造 32MB：直接验证错误话术）
        let error = too_large_media("word/media/big.png", 40 * 1024 * 1024);
        assert!(error.contains("太大"), "{error}");
        assert!(error.contains("40.0 MB"), "要说清多大：{error}");
        assert!(error.contains("Word / WPS"), "要给出替代办法：{error}");
        assert_eq!(human_size(512), "512 字节");
        assert_eq!(human_size(2048), "2 KB");
        assert_eq!(human_size(3 * 1024 * 1024), "3.0 MB");
    }

    #[test]
    fn media_mime_covers_common_word_formats() {
        for (name, mime) in [
            ("word/media/a.png", "image/png"),
            ("word/media/a.JPG", "image/jpeg"),
            ("word/media/a.jpeg", "image/jpeg"),
            ("word/media/a.gif", "image/gif"),
            ("word/media/a.bmp", "image/bmp"),
            ("word/media/a.webp", "image/webp"),
            ("word/media/a.svg", "image/svg+xml"),
            ("word/media/a.tif", "image/tiff"),
            ("word/media/a.tiff", "image/tiff"),
            ("word/media/a.ico", "image/x-icon"),
            ("word/media/a.emf", "image/emf"),
            ("word/media/a.wmf", "image/wmf"),
            ("word/media/a.bin", "application/octet-stream"),
            ("word/media/noext", "application/octet-stream"),
        ] {
            assert_eq!(media_mime(name), mime, "{name}");
        }
    }

    #[test]
    fn document_media_reads_real_sample_png_if_present() {
        let sample = std::env::var("MASTEREDIT_DOCX_SAMPLE")
            .unwrap_or_else(|_| r"C:\Users\master\Documents\7.20-24周报.docx".to_string());
        if !Path::new(&sample).exists() {
            eprintln!("跳过：样本不存在 {sample}");
            return;
        }

        // 1) 块模型给出的 media 名必须能直接喂给 document_media（这条把两个命令接起来了：
        //    也验证了 rels 里的 Target 归一化后的名字与 zip 里的部件名一致）
        let page = document_blocks(sample.clone(), 0, MAX_BLOCK_WINDOW).expect("取块应成功");
        let media: Vec<String> = page
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Image(image) => Some(image.media.clone()),
                _ => None,
            })
            .collect();
        assert_eq!(media.len(), 2, "样本里两张图");
        for name in &media {
            let url = document_media(sample.clone(), name.clone())
                .unwrap_or_else(|error| panic!("{name} 应能取到：{error}"));
            let payload = url
                .strip_prefix("data:image/png;base64,")
                .unwrap_or_else(|| panic!("{name} 应是 PNG"));
            let decoded = file::base64_decode(payload).expect("应能解码");
            assert_eq!(decoded[0..8], PNG_MAGIC, "{name} 应是 PNG 魔数");
            assert!(decoded.len() > 1024, "{name} 不该是空图");
        }

        // 2) 顺带量一下"每张图都重新解密 + 解包"的代价（前端是按需逐张取的）
        let start = std::time::Instant::now();
        for _ in 0..5 {
            document_media(sample.clone(), media[0].clone()).expect("应能取到");
        }
        let elapsed = start.elapsed();
        eprintln!(
            "真实样本取图：5 次共 {:?}（平均 {:?}/张）",
            elapsed,
            elapsed / 5
        );
    }

    /// **给前端的 vMerge 契约**（前端算 rowspan 就按这套走）：
    /// `vMerge = "continue"` 的单元格**仍然出现在 `rows[i].cells` 里，不会被省略**，
    /// 所以前端按"本行网格列"（前面单元格的 `gridSpan` 累加）往上找最近的 `restart` 即可配对。
    /// 这个测试就用前端将要写的那套逻辑算一遍 rowspan，证明契约够用。
    #[test]
    fn frontend_can_compute_rowspan_from_v_merge_contract() {
        let body = r#"
            <w:tbl>
              <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
              <w:tr>
                <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
                <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
                <w:tc><w:p><w:r><w:t>F</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
              </w:tr>
            </w:tbl>"#;
        let blocks = blocks_of(&docx_with(body, &[]));
        let table = table_of(&blocks[0]);

        // 前端第一步：把每个单元格铺到网格列上（(起始列, 跨列数, 纵向合并状态)）
        let mut layout: Vec<Vec<(usize, usize, VMerge)>> = Vec::new();
        for row in &table.rows {
            let mut column = 0usize;
            let mut cells = Vec::new();
            for cell in &row.cells {
                cells.push((column, cell.grid_span, cell.v_merge));
                column += cell.grid_span;
            }
            layout.push(cells);
        }

        // continue 单元格必须都在（否则前端无法按列配对，表格会错列）
        assert_eq!(
            layout,
            vec![
                vec![
                    (0, 1, VMerge::Restart),
                    (1, 1, VMerge::None),
                    (2, 1, VMerge::None)
                ],
                vec![
                    (0, 1, VMerge::Continue),
                    (1, 1, VMerge::None),
                    (2, 1, VMerge::Restart)
                ],
                vec![
                    (0, 1, VMerge::Continue),
                    (1, 1, VMerge::None),
                    (2, 1, VMerge::Continue)
                ],
            ]
        );

        // 前端第二步：每个 restart 往下数连续的 continue → rowspan
        let mut rowspans: Vec<((usize, usize), usize)> = Vec::new();
        for (row_index, row) in layout.iter().enumerate() {
            for &(column, span, merge) in row {
                if merge != VMerge::Restart {
                    continue;
                }
                let mut rowspan = 1usize;
                let mut probe = row_index + 1;
                while probe < layout.len() {
                    let continued = layout[probe]
                        .iter()
                        .any(|&(cell_column, cell_span, cell_merge)| {
                            cell_column == column
                                && cell_span == span
                                && cell_merge == VMerge::Continue
                        });
                    if !continued {
                        break;
                    }
                    rowspan += 1;
                    probe += 1;
                }
                rowspans.push(((row_index, column), rowspan));
            }
        }
        assert_eq!(
            rowspans,
            vec![((0, 0), 3), ((1, 2), 2)],
            "第一列竖跨 3 行；第三列从第 2 行起竖跨 2 行"
        );
    }

    /// **字符单位的缩进**（只写 `*Chars` 的生成器不能丢首行缩进）：
    /// 1 字符 = 本段**最终生效的字号**；显式 twips 存在时以 twips 为准。
    #[test]
    fn resolves_character_unit_indents_against_effective_font_size() {
        // docDefaults 字号 21 半磅（10.5pt）；样式 Big 给 48 半磅（24pt）——用来证明
        // 折算用的是"样式链算完之后"的字号，而不是某个固定值
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="Big"><w:name w:val="Big"/><w:basedOn w:val="Normal"/>
                <w:rPr><w:sz w:val="48"/></w:rPr></w:style>
            </w:styles>"#
        );
        let body = r#"
            <w:p><w:pPr><w:ind w:firstLineChars="200"/></w:pPr><w:r><w:t>两字符首行</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Big"/><w:ind w:firstLineChars="200"/></w:pPr>
              <w:r><w:t>大字两字符</w:t></w:r></w:p>
            <w:p><w:pPr><w:ind w:firstLineChars="200" w:firstLine="999"/></w:pPr>
              <w:r><w:t>两者都写</w:t></w:r></w:p>
            <w:p><w:pPr><w:ind w:hangingChars="200"/></w:pPr><w:r><w:t>悬挂两字符</w:t></w:r></w:p>
            <w:p><w:pPr><w:ind w:leftChars="150" w:rightChars="50"/></w:pPr>
              <w:r><w:t>左右字符</w:t></w:r></w:p>
            <w:p><w:r><w:t>没有缩进</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(
            body,
            &[("word/styles.xml", &styles)],
        ));

        // 1) 只写 Chars：10.5pt × 2 字符 = 21pt
        let first = paragraph_of(&blocks[0]);
        assert_eq!(first.runs[0].size_pt, Some(10.5), "字号来自 docDefaults");
        assert_eq!(
            first.indent_first_line_pt,
            Some(21.0),
            "firstLineChars=200 → 2 字符 × 10.5pt"
        );
        assert_eq!(first.indent_left_pt, None, "没写 leftChars 就不要凭空造一个");

        // 2) 字号来自样式链：24pt × 2 = 48pt（证明用的是最终生效字号）
        let big = paragraph_of(&blocks[1]);
        assert_eq!(big.runs[0].size_pt, Some(24.0));
        assert_eq!(big.indent_first_line_pt, Some(48.0), "同 200 Chars，字号变了值就变");

        // 3) 同时有 twips 与 Chars：以 twips 为准（Word 写文档时两者都写，观感按 twips）
        let both = paragraph_of(&blocks[2]);
        assert!(
            (both.indent_first_line_pt.unwrap() - 49.95).abs() < 0.01,
            "999 twips = 49.95pt，实际 {:?}",
            both.indent_first_line_pt
        );

        // 4) hangingChars → 负的首行缩进
        assert_eq!(paragraph_of(&blocks[3]).indent_first_line_pt, Some(-21.0));

        // 5) leftChars / rightChars
        let sides = paragraph_of(&blocks[4]);
        assert_eq!(sides.indent_left_pt, Some(15.75), "1.5 字符 × 10.5pt");
        assert_eq!(sides.indent_right_pt, Some(5.25), "0.5 字符 × 10.5pt");

        // 6) 普通段落不凭空有缩进
        let plain = paragraph_of(&blocks[5]);
        assert_eq!(plain.indent_left_pt, None);
        assert_eq!(plain.indent_first_line_pt, None);
    }

    /// 字符单位缩进在**编号级别**上也要生效（有的生成器把「2 字符」写在 numbering.xml 里）
    #[test]
    fn resolves_character_unit_indents_from_numbering_level() {
        let numbering = format!(
            r#"<?xml version="1.0"?><w:numbering {NS}>
              <w:abstractNum w:abstractNumId="0">
                <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
                  <w:pPr><w:ind w:leftChars="200" w:hangingChars="200"/></w:pPr></w:lvl>
              </w:abstractNum>
              <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
            </w:numbering>"#
        );
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
            </w:styles>"#
        );
        let body = r#"<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>
            <w:r><w:t>列表项</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(
            body,
            &[
                ("word/numbering.xml", &numbering),
                ("word/styles.xml", &styles),
            ],
        ));
        let paragraph = paragraph_of(&blocks[0]);
        assert_eq!(paragraph.runs[0].size_pt, Some(12.0), "docDefaults 24 半磅 = 12pt");
        assert_eq!(
            paragraph.indent_left_pt,
            Some(24.0),
            "编号级别的 leftChars=200 → 2 字符 × 12pt"
        );
        assert_eq!(paragraph.indent_first_line_pt, Some(-24.0), "hangingChars → 负值");
    }

    /// **自动段间距**（`w:beforeAutospacing` / `w:afterAutospacing`）：
    /// Word 显示为 Auto，约 14pt；docDefaults 里的 `before="0"` 不能把它吃掉；
    /// 段落自己写了显式值时以显式值为准。
    #[test]
    fn applies_automatic_paragraph_spacing() {
        // docDefaults 明确写了 before=0/after=0（现实文档很常见），用来验证 Auto 不被它盖掉
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault>
                <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault>
              </w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="Auto"><w:name w:val="Auto"/>
                <w:pPr><w:spacing w:beforeAutospacing="1"/></w:pPr></w:style>
            </w:styles>"#
        );
        let body = r#"
            <w:p><w:pPr><w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr>
              <w:r><w:t>自动间距</w:t></w:r></w:p>
            <w:p><w:pPr><w:spacing w:before="120" w:beforeAutospacing="1"/></w:pPr>
              <w:r><w:t>显式覆盖</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Auto"/></w:pPr><w:r><w:t>样式里的自动</w:t></w:r></w:p>
            <w:p><w:r><w:t>普通段落</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[("word/styles.xml", &styles)]));

        // 1) 段落自己开了自动：段前段后都是 Auto（≈14pt），docDefaults 的 0 不算数
        let auto = paragraph_of(&blocks[0]);
        assert_eq!(auto.space_before_pt, Some(AUTO_SPACE_PT));
        assert_eq!(auto.space_after_pt, Some(AUTO_SPACE_PT));

        // 2) 显式 before="120"（6pt）压过自动；after 没开自动 → 仍取 docDefaults 的 0
        let explicit = paragraph_of(&blocks[1]);
        assert_eq!(explicit.space_before_pt, Some(6.0), "显式 twips 优先");
        assert_eq!(explicit.space_after_pt, Some(0.0), "没开自动就按 docDefaults");

        // 3) 自动开关写在样式里一样生效（层叠带过来）
        let from_style = paragraph_of(&blocks[2]);
        assert_eq!(from_style.space_before_pt, Some(AUTO_SPACE_PT));
        assert_eq!(from_style.space_after_pt, Some(0.0));

        // 4) 没开自动的段落不受影响
        let plain = paragraph_of(&blocks[3]);
        assert_eq!(plain.space_before_pt, Some(0.0));
        assert_eq!(plain.space_after_pt, Some(0.0));
    }

    /// **行单位的段间距**（`w:beforeLines` / `w:afterLines`，百分之一行）：Word/WPS 会跟
    /// twips 一起写，但别的生成器可能只写它 —— 不折算段前/段后就会偏紧。
    /// 折算口径：1 行 = 该段最终行距折出的单行高度（见 `line_height_pt`）。
    #[test]
    fn resolves_line_unit_paragraph_spacing() {
        // 样式：docDefaults 12pt 字号、**不给行距**（用单倍行距基准 = 12 × 1.2 = 14.4pt）
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>
                <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault>
              </w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            </w:styles>"#
        );
        let body = r#"
            <w:p><w:pPr><w:spacing w:beforeLines="200"/></w:pPr><w:r><w:t>两行段前</w:t></w:r></w:p>
            <w:p><w:pPr><w:spacing w:beforeLines="200" w:before="120"/></w:pPr>
              <w:r><w:t>twips 优先</w:t></w:r></w:p>
            <w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact" w:afterLines="150"/></w:pPr>
              <w:r><w:t>固定行距 20pt</w:t></w:r></w:p>
            <w:p><w:pPr><w:spacing w:line="480" w:lineRule="auto" w:beforeLines="100"/></w:pPr>
              <w:r><w:t>两倍行距</w:t></w:r></w:p>
            <w:p><w:r><w:t>什么都没写</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[("word/styles.xml", &styles)]));

        // ① 只写 beforeLines=200、字号 12pt、无行距 → 2 × (12 × 1.2) = 28.8
        let two_lines = paragraph_of(&blocks[0]);
        assert_eq!(two_lines.runs[0].size_pt, Some(12.0));
        assert!(
            (two_lines.space_before_pt.unwrap() - 28.8).abs() < 0.001,
            "2 行 × 14.4pt = 28.8，实际 {:?}",
            two_lines.space_before_pt
        );

        // ② 同时有 before="120"（6pt）→ twips 优先
        let explicit = paragraph_of(&blocks[1]);
        assert_eq!(explicit.space_before_pt, Some(6.0), "段落显式 twips 压过行单位");

        // ③ 固定行距 20pt + afterLines=150 → 1.5 × 20 = 30.0
        let exact = paragraph_of(&blocks[2]);
        assert_eq!(exact.line_spacing.as_ref().map(|line| line.kind.as_str()), Some("exact"));
        assert!(
            (exact.space_after_pt.unwrap() - 30.0).abs() < 0.001,
            "1.5 行 × 20pt = 30.0，实际 {:?}",
            exact.space_after_pt
        );

        // ④ 两倍行距（line=480 auto）+ beforeLines=100 → 1 × (12 × 1.2 × 2) = 28.8
        let double = paragraph_of(&blocks[3]);
        assert!(
            (double.space_before_pt.unwrap() - 28.8).abs() < 0.001,
            "1 行 × 28.8pt = 28.8，实际 {:?}",
            double.space_before_pt
        );

        // ⑤ 什么都没写 → 沿用层叠值（docDefaults 的 0），不要变 null
        let plain = paragraph_of(&blocks[4]);
        assert_eq!(plain.space_before_pt, Some(0.0));
        assert_eq!(plain.space_after_pt, Some(0.0));
    }

    /// 行单位与自动段间距同时出现时的优先级：显式 twips > 行单位 > Auto > 层叠
    #[test]
    fn line_units_outrank_autospacing() {
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>
                <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault>
              </w:docDefaults>
            </w:styles>"#
        );
        let body = r#"
            <w:p><w:pPr><w:spacing w:beforeLines="100" w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr>
              <w:r><w:t>行单位压过自动</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[("word/styles.xml", &styles)]));
        let paragraph = paragraph_of(&blocks[0]);
        assert!(
            (paragraph.space_before_pt.unwrap() - 14.4).abs() < 0.001,
            "段前按 1 行 = 14.4pt（行单位优先），实际 {:?}",
            paragraph.space_before_pt
        );
        assert_eq!(
            paragraph.space_after_pt,
            Some(AUTO_SPACE_PT),
            "段后没写行单位 → 落到 Auto（14pt）"
        );
    }

    /// **层优先**：OOXML 里直接格式本就压过样式，与"用 twips 还是 Chars/Lines 写"无关。
    /// 样式里给 twips、段落里只给 Chars/Lines 时，按**段落**的写法算。
    #[test]
    fn paragraph_units_outrank_style_twips() {
        // 字号 12pt（docDefaults sz=24）；样式 StyleTwips 给 twips 缩进与段前，
        // 样式 StyleChars 给 Chars 缩进 —— 都挑与段落折算结果不同的数值，便于区分
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>
                <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault>
              </w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="StyleTwips"><w:name w:val="StyleTwips"/>
                <w:basedOn w:val="Normal"/>
                <w:pPr><w:ind w:left="720" w:firstLine="240"/><w:spacing w:before="240"/></w:pPr></w:style>
              <w:style w:type="paragraph" w:styleId="StyleChars"><w:name w:val="StyleChars"/>
                <w:basedOn w:val="Normal"/>
                <w:pPr><w:ind w:firstLineChars="300"/></w:pPr></w:style>
            </w:styles>"#
        );
        let body = r#"
            <w:p><w:pPr><w:pStyle w:val="StyleTwips"/>
                <w:ind w:firstLineChars="200" w:leftChars="100"/></w:pPr>
              <w:r><w:t>段落 Chars 压过样式 twips</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="StyleTwips"/><w:spacing w:beforeLines="100"/></w:pPr>
              <w:r><w:t>段落 Lines 压过样式 twips</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="StyleChars"/><w:ind w:firstLine="360"/></w:pPr>
              <w:r><w:t>段落 twips 压过样式 Chars</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="StyleTwips"/></w:pPr>
              <w:r><w:t>段落没写就用样式</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[("word/styles.xml", &styles)]));

        // 1) 段落 Chars（1 字符 / 2 字符 × 12pt）压过样式 twips（left=720→36pt、firstLine=240→12pt）
        let chars = paragraph_of(&blocks[0]);
        assert_eq!(chars.runs[0].size_pt, Some(12.0));
        assert_eq!(chars.indent_left_pt, Some(12.0), "1 字符 × 12pt（不是样式的 36pt）");
        assert_eq!(
            chars.indent_first_line_pt,
            Some(24.0),
            "2 字符 × 12pt（不是样式的 12pt）"
        );

        // 2) 段落 Lines（1 行 = 12 × 1.2 = 14.4pt）压过样式的 before="240"（12pt）
        let lines = paragraph_of(&blocks[1]);
        assert!(
            (lines.space_before_pt.unwrap() - 14.4).abs() < 0.001,
            "1 行 = 14.4pt（不是样式的 12pt），实际 {:?}",
            lines.space_before_pt
        );

        // 3) 反方向：段落 twips（360 twips = 18pt）压过样式的 firstLineChars="300"（36pt）
        let twips = paragraph_of(&blocks[2]);
        assert_eq!(twips.indent_first_line_pt, Some(18.0), "段落的 twips 说了算");

        // 4) 段落什么都没写 → 用样式的 twips
        let from_style = paragraph_of(&blocks[3]);
        assert_eq!(from_style.indent_left_pt, Some(36.0));
        assert_eq!(from_style.indent_first_line_pt, Some(12.0));
        assert_eq!(from_style.space_before_pt, Some(12.0));
    }

    /* ------------------------------ 页面几何（一页一页显示用） ------------------------------ */

    /// A4 竖版：`w:pgSz` 595.3×841.9pt（11906×16838 twip）、`w:pgMar` 72pt（1440 twip）
    #[test]
    fn parses_a4_portrait_page_geometry() {
        let body = r#"<w:p><w:r><w:t>正文</w:t></w:r></w:p>
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
              <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>"#;
        let parsed = parse(&docx_with(body, &[]));
        let page = parsed.page.clone().expect("应有页面几何（w:body/w:sectPr）");

        assert!((page.width_pt - 595.3).abs() < 0.01, "{}", page.width_pt);
        assert!((page.height_pt - 841.9).abs() < 0.01, "{}", page.height_pt);
        assert_eq!(page.margin_top_pt, 72.0);
        assert_eq!(page.margin_right_pt, 72.0);
        assert_eq!(page.margin_bottom_pt, 72.0);
        assert_eq!(page.margin_left_pt, 72.0);
        assert!(!page.landscape);
        // header/footer/gutter 不导出（前端分页用不到）

        // 前端契约：字段名 camelCase，page 挂在 BlockPage 上
        let page_block = BlockPage {
            total: parsed.blocks.len(),
            from: 0,
            blocks: parsed.blocks.clone(),
            encrypted: false,
            page: parsed.page.clone(),
            header: parsed.header.clone(),
            footer: parsed.footer.clone(),
        };
        let json = serde_json::to_value(&page_block).expect("应能序列化");
        assert!((json["page"]["widthPt"].as_f64().unwrap() - 595.3).abs() < 0.01);
        assert!((json["page"]["heightPt"].as_f64().unwrap() - 841.9).abs() < 0.01);
        assert_eq!(json["page"]["marginTopPt"], 72.0);
        assert_eq!(json["page"]["landscape"], false);
    }

    /// 横版：Word 会把 `w:w/w:h` 直接写成横向尺寸，同时给 `w:orient="landscape"`
    #[test]
    fn parses_landscape_page_geometry() {
        let body = r#"<w:p><w:r><w:t>横向</w:t></w:r></w:p>
            <w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
              <w:pgMar w:top="1800" w:right="1440" w:bottom="1800" w:left="1440"/></w:sectPr>"#;
        let page = parse(&docx_with(body, &[])).page.expect("应有页面几何");
        assert!(page.landscape);
        assert!((page.width_pt - 841.9).abs() < 0.01, "{}", page.width_pt);
        assert!((page.height_pt - 595.3).abs() < 0.01, "{}", page.height_pt);
        assert_eq!(page.margin_top_pt, 90.0, "1800 twip = 90pt");
        assert_eq!(page.margin_left_pt, 72.0);
    }

    /// 没有 `w:sectPr` → `page` 为 `None`（前端按 A4 + 2.54cm 兜底）
    /// 退化路径：`w:sectPr` 藏在段落属性里（`w:pPr/w:sectPr`，分节符的写法）
    #[test]
    fn page_geometry_falls_back_to_paragraph_sectpr_or_none() {
        // 1) 完全没有 sectPr
        let none = parse(&docx_with(r#"<w:p><w:r><w:t>无页面设置</w:t></w:r></w:p>"#, &[]));
        assert!(none.page.is_none(), "没有 sectPr 就该是 None");

        // 2) 只有段落属性里的 sectPr（且取最后一个）
        let body = r#"
            <w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
              <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr>
              <w:r><w:t>第一节</w:t></w:r></w:p>
            <w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
              <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:pPr>
              <w:r><w:t>第二节</w:t></w:r></w:p>"#;
        let page = parse(&docx_with(body, &[]))
            .page
            .expect("段落里的 sectPr 也要认");
        assert!(page.landscape, "取最后一个（第二节）");
        assert_eq!(page.margin_top_pt, 36.0, "720 twip = 36pt");
    }

    /// 畸形值不 panic：非数字、0、缺属性、只有部分属性 —— 单字段按 A4 兜底
    #[test]
    fn malformed_page_geometry_values_fall_back_to_a4() {
        let cases = [
            // 非数字 + 缺 h
            r#"<w:pgSz w:w="abc"/><w:pgMar w:top="xyz" w:right="1440"/>"#,
            // 0 尺寸（Word 里不该出现，但坏文档会有）+ 负边距（Word 允许）
            r#"<w:pgSz w:w="0" w:h="-5"/><w:pgMar w:top="-100" w:right="0" w:bottom="1440" w:left="720"/>"#,
            // 空 sectPr
            r#""#,
            // 只有 pgMar
            r#"<w:pgMar w:top="1440"/>"#,
        ];
        for case in cases {
            let body = format!("<w:p><w:r><w:t>坏设置</w:t></w:r></w:p><w:sectPr>{case}</w:sectPr>");
            let parsed = parse(&docx_with(&body, &[]));
            let page = parsed.page.unwrap_or_else(|| panic!("{case} 应仍给出页面几何"));
            assert!(page.width_pt > 0.0 && page.height_pt > 0.0, "{case}");
            assert!(
                page.margin_top_pt.is_finite() && page.margin_left_pt.is_finite(),
                "{case}"
            );
        }

        // 逐个字段确认兜底值
        let page = parse(&docx_with(
            r#"<w:p/><w:sectPr><w:pgSz w:w="abc"/><w:pgMar w:top="xyz"/></w:sectPr>"#,
            &[],
        ))
        .page
        .expect("应给出兜底几何");
        assert_eq!(page.width_pt, A4_WIDTH_PT, "非数字 → A4 宽");
        assert_eq!(page.height_pt, A4_HEIGHT_PT, "缺属性 → A4 高");
        assert_eq!(page.margin_top_pt, A4_MARGIN_PT, "非数字 → A4 边距");
        assert_eq!(page.margin_right_pt, A4_MARGIN_PT);
        assert!(!page.landscape);

        // 0 尺寸 → A4；负边距 / 0 边距要**原样保留**（Word 允许）
        let page = parse(&docx_with(
            r#"<w:p/><w:sectPr><w:pgSz w:w="0" w:h="0"/>
                 <w:pgMar w:top="-100" w:right="0" w:bottom="1440" w:left="720"/></w:sectPr>"#,
            &[],
        ))
        .page
        .expect("应给出兜底几何");
        assert_eq!(page.width_pt, A4_WIDTH_PT, "0 宽 → A4");
        assert_eq!(page.height_pt, A4_HEIGHT_PT, "0 高 → A4");
        assert_eq!(page.margin_top_pt, -5.0, "负边距保留（-100 twip）");
        assert_eq!(page.margin_right_pt, 0.0, "0 边距保留");
        assert_eq!(page.margin_bottom_pt, 72.0);
    }

    /// **形状渲染**（合同/表单里的横线与方框）：VML 与 DrawingML 两条路都要覆盖。
    /// 认不出来的（组合图形、无 `prstGeom`、旋转过的）仍然走占位块 —— **不丢内容**。
    #[test]
    fn renders_vml_and_drawingml_shapes() {
        let rels = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/vml.png"/>
            </Relationships>"#
        );
        // graphicData 的 uri 决定走图片/图形/SmartArt/图表哪条路，测试里写真实的 uri。
        // `wp:extent` 是 Word 排版用的显示尺寸（也是我们优先取的那个），所以与 `a:ext` 保持一致。
        let drawing_full = |uri: &str, cx: &str, cy: &str, body: &str| {
            format!(
                r#"<w:p><w:r><w:drawing><wp:inline><wp:extent cx="{cx}" cy="{cy}"/>
                  <a:graphic><a:graphicData uri="{uri}">
                    {body}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
            )
        };
        let shape_uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        let drawing = |body: &str| drawing_full(shape_uri, "914400", "0", body);
        let body = format!(
            r#"{vml_line}{vml_hr}{vml_rect}{dml_line}{dml_rect}{alternate}{vml_two_lines}
               {vml_group}{dml_no_prst}{dml_rotated}{diagram}{chart}{missing_image}"#,
            // ① VML 直线（合同里的横线）：400pt 长、height:0 + strokeweight 1pt
            vml_line = r##"<w:p><w:r><w:pict><v:line from="0,0" to="400pt,0"
                style="position:absolute;left:0;text-align:left;width:400pt;height:0"
                strokecolor="#3465a4" strokeweight="1pt"/></w:pict></w:r></w:p>"##,
            // ② Word 的「水平线」：没有尺寸，就是整栏宽
            vml_hr = r#"<w:p><w:r><w:pict><v:hr o:hrpct="100" o:hrstd="t"/></w:pict></w:r></w:p>"#,
            // ③ VML 矩形（表单方框）：无填充 + 具名颜色 black
            vml_rect = r#"<w:p><w:r><w:pict><v:rect style="width:200pt;height:40pt"
                filled="f" strokecolor="black" strokeweight=".5pt"/></w:pict></w:r></w:p>"#,
            // ④ DrawingML 直线：5486400 EMU = 432pt，1pt 红色虚线
            dml_line = drawing_full(
                shape_uri,
                "5486400",
                "0",
                r#"<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="0"/></a:xfrm>
                   <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
                   <a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                     <a:prstDash val="dash"/></a:ln>
                   </wps:spPr></wps:wsp>"#
            ),
            // ⑤ DrawingML 矩形：914400×457200 EMU = 72×36pt，带填充
            dml_rect = drawing_full(
                shape_uri,
                "914400",
                "457200",
                r#"<wps:wsp><wps:spPr><a:xfrm><a:off x="12700" y="25400"/><a:ext cx="914400" cy="457200"/></a:xfrm>
                   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                   <a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>
                   <a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                   </wps:spPr></wps:wsp>"#
            ),
            // ⑥ mc:AlternateContent：Choice=DrawingML、Fallback=VML → 只能产出一个 Shape
            alternate = r#"<w:p><w:r><mc:AlternateContent>
                <mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>
                  <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp>
                  </a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>
                <mc:Fallback><w:pict><v:rect style="width:100pt;height:20pt"/></w:pict></mc:Fallback>
                </mc:AlternateContent></w:r></w:p>"#,
            // ⑦ 一个 w:pict 里两条线 → 两个 Shape 块（不是只认第一个）
            vml_two_lines = r#"<w:p><w:r><w:pict>
                <v:line from="0,0" to="100pt,0" style="width:100pt;height:0" strokeweight="1pt"/>
                <v:line from="0,0" to="0,50pt" style="width:0;height:50pt" strokeweight="1pt"/>
                </w:pict></w:r></w:p>"#,
            // ⑧ 组合图形：整体走占位（只画其中一部分会丢内容）
            vml_group = r#"<w:p><w:r><w:pict><v:group style="width:100pt;height:50pt">
                <v:rect style="width:100pt;height:50pt"/><v:line from="0,0" to="100pt,0"/></v:group></w:pict></w:r></w:p>"#,
            // ⑨ DrawingML 自由曲线（没有 prstGeom）→ 占位
            dml_no_prst = drawing(
                r#"<wps:wsp><wps:spPr><a:xfrm><a:ext cx="914400" cy="914400"/></a:xfrm>
                   <a:custGeom><a:pathLst/></a:custGeom></wps:spPr></wps:wsp>"#
            ),
            // ⑩ 旋转过的形状 → 占位（几何算不准，画歪不如不画）
            dml_rotated = drawing(
                r#"<wps:wsp><wps:spPr><a:xfrm rot="5400000"><a:ext cx="914400" cy="457200"/></a:xfrm>
                   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp>"#
            ),
            diagram = drawing_full(
                "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                "914400",
                "914400",
                r#"<dgm:relIds r:dm="rId1"/>"#
            ),
            chart = drawing_full(
                "http://schemas.openxmlformats.org/drawingml/2006/chart",
                "914400",
                "914400",
                r#"<c:chart r:id="rId2"/>"#
            ),
            missing_image = drawing(
                r#"<pic:pic><pic:blipFill><a:blip r:embed="rId404"/></pic:blipFill></pic:pic>"#
            ),
        );
        let blocks = blocks_of(&docx_with(&body, &[("word/_rels/document.xml.rels", &rels)]));

        let kinds: Vec<String> = blocks
            .iter()
            .map(|block| match block {
                Block::Shape(shape) => format!("形状:{}", shape.shape),
                Block::Unsupported { label, .. } => format!("占位:{label}"),
                other => format!("{other:?}"),
            })
            .collect();

        // ① VML 直线：400pt 长；height 为 0 → 按 strokeweight（1pt）兜底，前端才好定位
        let line = match &blocks[0] {
            Block::Shape(shape) => shape,
            other => panic!("VML 直线应产出 Shape：{other:?}"),
        };
        assert_eq!(line.shape, "line");
        assert_eq!(line.width_pt, 400.0);
        assert_eq!(line.height_pt, 1.0, "height:0 → 按线宽 1pt 兜底（不是 0）");
        assert_eq!(line.line_width_pt, Some(1.0));
        assert_eq!(line.line_color.as_deref(), Some("3465A4"));
        assert_eq!(line.fill_color, None);
        assert_eq!(line.dash, None);
        assert!(!line.vertical, "横线");

        // ② v:hr：整栏宽（无 sectPr → A4 正文宽 595.3 − 72×2 = 451.3pt）
        let hr = match &blocks[1] {
            Block::Shape(shape) => shape,
            other => panic!("v:hr 应产出 Shape：{other:?}"),
        };
        assert_eq!(hr.shape, "line");
        assert!((hr.width_pt - 451.3).abs() < 0.1, "实际 {}", hr.width_pt);
        assert_eq!(hr.height_pt, 1.0);
        assert_eq!(hr.line_color, None, "拿不到颜色给 null（前端用默认黑）");

        // ③ VML 矩形：具名颜色 black → 000000；filled="f" → 无填充
        let rect = match &blocks[2] {
            Block::Shape(shape) => shape,
            other => panic!("VML 矩形应产出 Shape：{other:?}"),
        };
        assert_eq!(rect.shape, "rect");
        assert_eq!((rect.width_pt, rect.height_pt), (200.0, 40.0));
        assert_eq!(rect.line_width_pt, Some(0.5));
        assert_eq!(rect.line_color.as_deref(), Some("000000"));
        assert_eq!(rect.fill_color, None);

        // ④ DrawingML 直线：432pt、1pt 红虚线、y=0
        let dml_line = match &blocks[3] {
            Block::Shape(shape) => shape,
            other => panic!("DrawingML 直线应产出 Shape：{other:?}"),
        };
        assert_eq!(dml_line.shape, "line");
        assert_eq!(dml_line.width_pt, 432.0);
        assert_eq!(dml_line.height_pt, 1.0, "cy=0 → 按 1pt 线宽兜底");
        assert_eq!(dml_line.line_width_pt, Some(1.0));
        assert_eq!(dml_line.line_color.as_deref(), Some("FF0000"));
        assert_eq!(dml_line.dash.as_deref(), Some("dash"));
        assert_eq!((dml_line.x_pt, dml_line.y_pt), (0.0, 0.0));

        // ⑤ DrawingML 矩形：72×36pt、填充 DDEEFF、2pt 黑边、偏移 (1pt, 2pt)
        let dml_rect = match &blocks[4] {
            Block::Shape(shape) => shape,
            other => panic!("DrawingML 矩形应产出 Shape：{other:?}"),
        };
        assert_eq!(dml_rect.shape, "rect");
        assert_eq!((dml_rect.width_pt, dml_rect.height_pt), (72.0, 36.0));
        assert_eq!(dml_rect.fill_color.as_deref(), Some("DDEEFF"));
        assert_eq!(dml_rect.line_width_pt, Some(2.0));
        assert_eq!(dml_rect.line_color.as_deref(), Some("000000"));
        assert_eq!((dml_rect.x_pt, dml_rect.y_pt), (1.0, 2.0));

        // ⑥ AlternateContent：只产出一个 Shape（Choice 的 DrawingML 那条，72×36）
        let alternate = match &blocks[5] {
            Block::Shape(shape) => shape,
            other => panic!("AlternateContent 应产出一个 Shape：{other:?}"),
        };
        assert_eq!((alternate.width_pt, alternate.height_pt), (72.0, 36.0));

        // ⑦ 一个 pict 两条线 → 两个 Shape：横线 + 竖线
        let two = match (&blocks[6], &blocks[7]) {
            (Block::Shape(first), Block::Shape(second)) => (first, second),
            other => panic!("一个 pict 里的两条线应产出两个 Shape：{other:?}"),
        };
        assert_eq!((two.0.width_pt, two.0.height_pt), (100.0, 1.0));
        assert!(!two.0.vertical);
        assert_eq!((two.1.width_pt, two.1.height_pt), (1.0, 50.0), "竖线宽度兜底 1pt");
        assert!(two.1.vertical, "height > width → 竖线");

        // ⑧⑨⑩ 认不出来的 → 占位，且说明里讲清是哪一类
        assert_eq!(kinds[8], "占位:图形对象（VML，暂不支持显示）");
        assert!(kinds[8].contains("VML"), "组合图形走 VML 占位");
        assert_eq!(kinds[9], "占位:图形对象（暂不支持显示）", "自由曲线");
        assert_eq!(kinds[10], "占位:图形对象（暂不支持显示）", "旋转过的形状");
        assert_eq!(kinds[11], "占位:SmartArt 图形（暂不支持显示）");
        assert_eq!(kinds[12], "占位:图表（暂不支持显示）");
        assert_eq!(kinds[13], "占位:图片（无法显示）");

        // 前端契约：camelCase 字段名 + kind 判别
        let json = serde_json::to_value(&blocks[4]).expect("应能序列化");
        assert_eq!(json["kind"], "shape");
        assert_eq!(json["shape"], "rect");
        assert_eq!(json["widthPt"], 72.0);
        assert_eq!(json["heightPt"], 36.0);
        assert_eq!(json["lineWidthPt"], 2.0);
        assert_eq!(json["lineColor"], "000000");
        assert_eq!(json["fillColor"], "DDEEFF");
        assert!(json["dash"].is_null());
        assert_eq!(json["vertical"], false);
    }

    /// `v:hr` 的"整栏宽"要跟着页面几何与段落缩进走
    #[test]
    fn hr_width_follows_page_and_indent() {
        let body = r#"
            <w:p><w:pPr><w:ind w:left="720" w:right="360"/></w:pPr>
              <w:r><w:pict><v:hr/></w:pict></w:r></w:p>
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
              <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#;
        let blocks = blocks_of(&docx_with(body, &[]));
        let hr = match &blocks[0] {
            Block::Shape(shape) => shape,
            other => panic!("v:hr 应产出 Shape：{other:?}"),
        };
        // 正文宽 595.3 − 72 − 72 = 451.3；再减左缩进 36pt（720 twip）与右缩进 18pt（360 twip）
        assert!(
            (hr.width_pt - (451.3 - 36.0 - 18.0)).abs() < 0.1,
            "实际 {}",
            hr.width_pt
        );
    }

    /// **段落边框**（`w:pBdr`）：合同/表单里"空段落 + 下边框"就是一条横线。
    /// 只解析 `w:val` 不为 none/nil 的边；`w:sz` 是八分之一磅；四边皆空 → 整个字段 null。
    #[test]
    fn parses_paragraph_borders() {
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="Boxed"><w:name w:val="Boxed"/>
                <w:pPr><w:pBdr>
                  <w:top w:val="double" w:sz="12" w:color="0000FF"/>
                  <w:left w:val="single" w:sz="8" w:color="auto"/>
                </w:pBdr></w:pPr></w:style>
            </w:styles>"#
        );
        let body = r#"
            <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="FF0000"/></w:pBdr></w:pPr>
              <w:r><w:t>只有下边框</w:t></w:r></w:p>
            <w:p><w:pPr><w:pBdr><w:bottom w:val="none" w:sz="6" w:color="FF0000"/></w:pBdr></w:pPr>
              <w:r><w:t>明确不要边框</w:t></w:r></w:p>
            <w:p><w:r><w:t>没有边框</w:t></w:r></w:p>
            <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="auto"/></w:pBdr></w:pPr>
              <w:r><w:t>自动色</w:t></w:r></w:p>
            <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="FFFFFF"/></w:pBdr></w:pPr>
              <w:r><w:t>白线</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Boxed"/></w:pPr><w:r><w:t>样式里的边框</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[("word/styles.xml", &styles)]));

        // ① 只写下边框：w:sz=6（八分之一磅）→ 0.75pt，其它三边 null
        let single = paragraph_of(&blocks[0])
            .borders
            .clone()
            .expect("应解析出段落边框");
        let bottom = single.bottom.as_ref().expect("应有下边框");
        assert_eq!(bottom.style, "single");
        assert!((bottom.width_pt - 0.75).abs() < 0.001, "{}", bottom.width_pt);
        assert_eq!(bottom.color.as_deref(), Some("FF0000"));
        assert_eq!(bottom.space_pt, Some(1.0));
        assert!(single.top.is_none() && single.left.is_none() && single.right.is_none());

        // ② w:val="none" 是"明确不要边框" → 四边皆空 → 整个字段 null
        assert!(paragraph_of(&blocks[1]).borders.is_none());
        // ③ 没有 w:pBdr → null
        assert!(paragraph_of(&blocks[2]).borders.is_none());

        // ④ w:color="auto" → null（前端用默认色）
        let auto = paragraph_of(&blocks[3]).borders.clone().expect("仍应有边框");
        assert_eq!(auto.bottom.expect("应有下边框").color, None);

        // ⑤ 白线要**保留**：白色是合法的可见设置，不能像底纹那样被当成"没写"
        let white = paragraph_of(&blocks[4]).borders.clone().expect("仍应有边框");
        assert_eq!(white.bottom.expect("应有下边框").color.as_deref(), Some("FFFFFF"));

        // ⑥ 样式里的边框也要能通过层叠拿到（top 是 double 1.5pt 蓝色、left 是 auto 色）
        let from_style = paragraph_of(&blocks[5]).borders.clone().expect("样式里的边框");
        let top = from_style.top.as_ref().expect("应有上边框");
        assert_eq!(top.style, "double");
        assert!((top.width_pt - 1.5).abs() < 0.001, "sz=12 → 1.5pt");
        assert_eq!(top.color.as_deref(), Some("0000FF"));
        assert_eq!(from_style.left.expect("应有左边框").color, None, "auto → null");
        assert!(from_style.bottom.is_none() && from_style.right.is_none());

        // 前端契约：camelCase 字段名
        let json = serde_json::to_value(&blocks[0]).expect("应能序列化");
        assert_eq!(json["borders"]["bottom"]["style"], "single");
        assert_eq!(json["borders"]["bottom"]["widthPt"], 0.75);
        assert_eq!(json["borders"]["bottom"]["color"], "FF0000");
        assert_eq!(json["borders"]["bottom"]["spacePt"], 1.0);
        assert!(json["borders"]["top"].is_null());
        let json = serde_json::to_value(&blocks[2]).expect("应能序列化");
        assert!(json["borders"].is_null(), "四边皆空 → null");
    }

    /// **多节文档**：一套页眉页脚取"离最后一节最近的那个 default"；最后一节没引用时往前找，
    /// 不能被"最后一节只有个页码"这种情况把整份文档的页眉丢掉（真实的企业标准就是这样）。
    #[test]
    fn header_footer_fall_back_to_earlier_sections() {
        let rels = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
              <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
              <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header4.xml"/>
              <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer6.xml"/>
            </Relationships>"#
        );
        let header2 = format!(
            r#"<?xml version="1.0"?><w:hdr {NS}><w:p><w:r><w:t>封面节的页眉</w:t></w:r></w:p></w:hdr>"#
        );
        let header4 = format!(
            r#"<?xml version="1.0"?><w:hdr {NS}><w:p><w:r><w:t>正文节的页眉</w:t></w:r></w:p></w:hdr>"#
        );
        let footer2 = format!(
            r#"<?xml version="1.0"?><w:ftr {NS}><w:p><w:r><w:t>封面节的页脚</w:t></w:r></w:p></w:ftr>"#
        );
        let footer6 = format!(r#"<?xml version="1.0"?><w:ftr {NS}><w:p><w:r><w:t>2</w:t></w:r></w:p></w:ftr>"#);
        let parts: Vec<(&str, &str)> = vec![
            ("word/_rels/document.xml.rels", &rels),
            ("word/header2.xml", &header2),
            ("word/header4.xml", &header4),
            ("word/footer2.xml", &footer2),
            ("word/footer6.xml", &footer6),
        ];

        // ① 最后一节（body 级）只有页脚 → 页眉往前找"离它最近"的 default（第一节）
        let body = r#"<w:p><w:r><w:t>封面</w:t></w:r></w:p>
            <w:p><w:pPr><w:sectPr>
              <w:headerReference w:type="default" r:id="rId8"/>
              <w:footerReference w:type="default" r:id="rId9"/>
            </w:sectPr></w:pPr><w:r><w:t>第一节结束</w:t></w:r></w:p>
            <w:p><w:r><w:t>正文</w:t></w:r></w:p>
            <w:sectPr><w:footerReference w:type="default" r:id="rId11"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>"#;
        let parsed = parse(&docx_with(body, &parts));
        let header = parsed.header.clone().expect("最后一节没页眉也要往前找到");
        assert_eq!(paragraph_of(&header[0]).text, "封面节的页眉");
        let footer = parsed.footer.clone().expect("最后一节的页脚");
        assert_eq!(paragraph_of(&footer[0]).text, "2");

        // ② 最后一节自己也有 default 页眉 → 用它（离最后一节最近的优先）
        let body = r#"<w:p><w:pPr><w:sectPr>
              <w:headerReference w:type="default" r:id="rId8"/>
            </w:sectPr></w:pPr><w:r><w:t>第一节结束</w:t></w:r></w:p>
            <w:sectPr><w:headerReference w:type="default" r:id="rId10"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>"#;
        let parsed = parse(&docx_with(body, &parts));
        let header = parsed.header.clone().expect("应有页眉");
        assert_eq!(paragraph_of(&header[0]).text, "正文节的页眉");

        // ③ 全都没有 default，只有 first → 仍然取得到（不被"没有 default"卡住）
        let body = r#"<w:p><w:pPr><w:sectPr>
              <w:headerReference w:type="first" r:id="rId8"/>
            </w:sectPr></w:pPr><w:r><w:t>第一节结束</w:t></w:r></w:p>
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>"#;
        let parsed = parse(&docx_with(body, &parts));
        let header = parsed.header.clone().expect("只有 first 也应取到");
        assert_eq!(paragraph_of(&header[0]).text, "封面节的页眉");
    }



    /// **文本框内容结构化**：企业标准封面整页、表单签字框都是文本框，不能只给一张占位卡片。
    #[test]
    fn renders_text_boxes_with_real_content() {
        let body = r##"
            <w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="228600"/>
              <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp><wps:spPr>
                    <a:xfrm><a:off x="12700" y="25400"/><a:ext cx="1828800" cy="228600"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>
                    <a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx><w:txbxContent>
                    <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Q/NDB 001—2026</w:t></w:r></w:p>
                    <w:p><w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>AI加持下的家庭能源管理系统</w:t></w:r></w:p>
                  </w:txbxContent></wps:txbx>
                </wps:wsp>
              </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
            <w:p><w:r><w:pict><v:shape style="width:200pt;height:40pt" fillcolor="#FFFFCC"
                strokecolor="black" strokeweight=".5pt">
              <v:textbox><w:txbxContent>
                <w:p><w:r><w:t>甲方签字：</w:t></w:r></w:p>
                <w:tbl><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
                  <w:tr><w:tc><w:p><w:r><w:t>表格也在文本框里</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
              </w:txbxContent></v:textbox>
            </v:shape></w:pict></w:r></w:p>
            <w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent/></v:textbox></v:shape></w:pict></w:r></w:p>
            <w:p><mc:AlternateContent>
              <mc:Choice Requires="wps"><w:pict><v:shape style="width:120pt;height:30pt">
                <v:textbox><w:txbxContent><w:p><w:r><w:t>Choice 里的框</w:t></w:r></w:p>
                </w:txbxContent></v:textbox></v:shape></w:pict></mc:Choice>
              <mc:Fallback><w:pict><v:shape style="width:120pt;height:30pt">
                <v:textbox><w:txbxContent><w:p><w:r><w:t>Fallback 里的框</w:t></w:r></w:p>
                </w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>
            </mc:AlternateContent></w:p>"##;
        let styles = format!(
            r#"<?xml version="1.0"?><w:styles {NS}>
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>
                <w:rPr><w:sz w:val="44"/><w:b/></w:rPr></w:style>
            </w:styles>"#
        );
        let blocks = blocks_of(&docx_with(body, &[("word/styles.xml", &styles)]));

        // ① DrawingML 文本框：内容结构化 + 尺寸/填充/边框/偏移
        let first = match &blocks[0] {
            Block::TextBox(text_box) => text_box,
            other => panic!("应产出文本框块：{other:?}"),
        };
        assert_eq!(first.blocks.len(), 2, "文本框里两个段落");
        assert_eq!(paragraph_of(&first.blocks[0]).text, "Q/NDB 001—2026");
        assert_eq!(paragraph_of(&first.blocks[1]).text, "AI加持下的家庭能源管理系统");
        // 文本框里的段落走同一套样式层叠：Title 样式给 22pt + 加粗
        let title_run = &paragraph_of(&first.blocks[0]).runs[0];
        assert_eq!(title_run.size_pt, Some(22.0), "sz=44 半磅来自 Title 样式");
        assert!(title_run.bold);
        assert_eq!(paragraph_of(&first.blocks[1]).runs[0].size_pt, Some(16.0), "sz=32 → 16pt");
        assert_eq!(paragraph_of(&first.blocks[1]).runs[0].bold, false);
        assert_eq!((first.width_pt, first.height_pt), (144.0, 18.0), "a:ext EMU → pt");
        assert_eq!((first.x_pt, first.y_pt), (1.0, 2.0), "a:off EMU → pt");
        assert_eq!(first.fill_color.as_deref(), Some("DDEEFF"));
        assert_eq!(first.border_color.as_deref(), Some("000000"));
        assert_eq!(first.border_width_pt, Some(2.0));
        assert_eq!(first.wrap, "none", "内联（wp:inline）→ none");

        // ② VML 文本框：尺寸来自 style；里面的表格照常解析
        let second = match &blocks[1] {
            Block::TextBox(text_box) => text_box,
            other => panic!("VML 文本框应产出文本框块：{other:?}"),
        };
        assert_eq!((second.width_pt, second.height_pt), (200.0, 40.0));
        assert_eq!(second.fill_color.as_deref(), Some("FFFFCC"));
        assert_eq!(second.border_color.as_deref(), Some("000000"));
        assert_eq!(second.border_width_pt, Some(0.5));
        assert_eq!(paragraph_of(&second.blocks[0]).text, "甲方签字：");
        let inner_table = table_of(&second.blocks[1]);
        assert_eq!(inner_table.rows[0].cells[0].text, "表格也在文本框里");

        // ③ 空文本框 → 仍然给占位（不静默丢）
        assert_eq!(label_of(&blocks[2]), "文本框（暂不支持显示）");

        // ④ AlternateContent：只出一个文本框（Choice 优先）
        let alternate = match &blocks[3] {
            Block::TextBox(text_box) => text_box,
            other => panic!("应产出一个文本框块：{other:?}"),
        };
        assert_eq!(paragraph_of(&alternate.blocks[0]).text, "Choice 里的框");
        assert_eq!(blocks.len(), 4, "不该两个都产出");

        // 前端契约：kind=textBox + camelCase 字段
        let json = serde_json::to_value(&blocks[0]).expect("应能序列化");
        assert_eq!(json["kind"], "textBox");
        assert_eq!(json["widthPt"], 144.0);
        assert_eq!(json["heightPt"], 18.0);
        assert_eq!(json["fillColor"], "DDEEFF");
        assert_eq!(json["borderColor"], "000000");
        assert_eq!(json["borderWidthPt"], 2.0);
        assert_eq!(json["wrap"], "none");
        assert_eq!(json["blocks"][0]["kind"], "paragraph");
        // find 索引里要能搜到封面文字
        let parsed = parse(&docx_with(body, &[("word/styles.xml", &styles)]));
        assert!(
            parsed.search_index.iter().any(|(_, text)| text.contains("Q/NDB 001—2026")),
            "文本框里的文字要能被搜索"
        );
    }

    /// 文本框是独立"故事"：里面的列表编号不该把正文的计数器往后推
    #[test]
    fn text_box_numbering_does_not_shift_body() {
        let numbering = format!(
            r#"<?xml version="1.0"?><w:numbering {NS}>
              <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
                <w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
              <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
        );
        let item = |text: &str| {
            format!(
                r#"<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
            )
        };
        let body = format!(
            r#"{before}<w:p><w:r><w:pict><v:shape style="width:100pt;height:30pt"><v:textbox>
                <w:txbxContent>{in_box}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>{after}"#,
            before = item("正文第一项"),
            in_box = item("框里第一项"),
            after = item("正文第二项"),
        );
        let blocks = blocks_of(&docx_with(&body, &[("word/numbering.xml", &numbering)]));

        let prefixes: Vec<String> = blocks
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => paragraph.list.as_ref().map(|list| list.prefix.clone()),
                _ => None,
            })
            .collect();
        assert_eq!(prefixes, vec!["1.", "2."], "文本框里的编号不吃掉正文的计数");

        let in_box = blocks
            .iter()
            .find_map(|block| match block {
                Block::TextBox(text_box) => Some(text_box),
                _ => None,
            })
            .expect("应有文本框块");
        let box_prefix = paragraph_of(&in_box.blocks[0])
            .list
            .as_ref()
            .map(|list| list.prefix.clone());
        assert_eq!(box_prefix.as_deref(), Some("1."), "框里自己从 1 开始");
    }

    /* ------------------------------ 页眉页脚 ------------------------------ */

    /// **文档自己的页眉页脚**：`w:headerReference` / `w:footerReference` → `word/header*.xml`。
    /// 只取 `w:type="default"`（没有 default 时取第一个），"首页不同"不区分。
    #[test]
    fn parses_header_and_footer_parts() {
        let rels = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
              <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
              <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
            </Relationships>"#
        );
        // 默认页眉 = header2.xml；首页页眉 = header1.xml（不该被取到）
        let header1 = format!(
            r#"<?xml version="1.0"?><w:hdr {NS}><w:p><w:r><w:t>首页页眉（不该用到）</w:t></w:r></w:p></w:hdr>"#
        );
        let header2 = format!(
            r#"<?xml version="1.0"?><w:hdr {NS}><w:p><w:r><w:t>AI加持下的家庭能源管理系统企业标准</w:t></w:r></w:p></w:hdr>"#
        );
        let footer1 = format!(
            r#"<?xml version="1.0"?><w:ftr {NS}><w:p><w:r><w:t>第 1 页 · 共 1 页</w:t></w:r></w:p></w:ftr>"#
        );
        let body = r#"<w:p><w:r><w:t>正文</w:t></w:r></w:p>
            <w:sectPr>
              <w:headerReference w:type="first" r:id="rId10"/>
              <w:headerReference w:type="default" r:id="rId8"/>
              <w:footerReference w:type="default" r:id="rId9"/>
              <w:pgSz w:w="11906" w:h="16838"/>
            </w:sectPr>"#;
        let parsed = parse(&docx_with(
            body,
            &[
                ("word/_rels/document.xml.rels", &rels),
                ("word/header1.xml", &header1),
                ("word/header2.xml", &header2),
                ("word/footer1.xml", &footer1),
            ],
        ));

        // ① default 页眉页脚都解析出来了，文本正确
        let header = parsed.header.clone().expect("应有页眉");
        assert_eq!(paragraph_of(&header[0]).text, "AI加持下的家庭能源管理系统企业标准");
        let footer = parsed.footer.clone().expect("应有页脚");
        assert_eq!(paragraph_of(&footer[0]).text, "第 1 页 · 共 1 页");

        // 窗口返回里也要带上（BlockPage.header/footer）
        let doc = TempDocx::new(
            "header",
            &docx_with(
                body,
                &[
                    ("word/_rels/document.xml.rels", &rels),
                    ("word/header1.xml", &header1),
                    ("word/header2.xml", &header2),
                    ("word/footer1.xml", &footer1),
                ],
            ),
        );
        let page = document_blocks(doc.path(), 0, 10).expect("取块应成功");
        assert_eq!(
            page.header.as_ref().map(|blocks| paragraph_of(&blocks[0]).text.clone()),
            Some("AI加持下的家庭能源管理系统企业标准".to_string())
        );
        assert_eq!(
            page.footer.as_ref().map(|blocks| paragraph_of(&blocks[0]).text.clone()),
            Some("第 1 页 · 共 1 页".to_string())
        );
        let json = serde_json::to_value(&page).expect("应能序列化");
        assert_eq!(json["header"][0]["kind"], "paragraph");
        assert_eq!(json["footer"][0]["text"], "第 1 页 · 共 1 页");
    }

    /// 页眉页脚的退化情形：只有 first 就取它；没有引用 / 引用缺失 / 畸形 id → `None`，不 panic
    #[test]
    fn header_footer_degrade_gracefully() {
        // ② 只有 first 页眉 → 取它
        let rels = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
            </Relationships>"#
        );
        let header1 = format!(
            r#"<?xml version="1.0"?><w:hdr {NS}><w:p><w:r><w:t>只有首页页眉</w:t></w:r></w:p></w:hdr>"#
        );
        let parsed = parse(&docx_with(
            r#"<w:p/><w:sectPr><w:headerReference w:type="first" r:id="rId10"/></w:sectPr>"#,
            &[
                ("word/_rels/document.xml.rels", &rels),
                ("word/header1.xml", &header1),
            ],
        ));
        let header = parsed.header.clone().expect("只有 first 也应取到");
        assert_eq!(paragraph_of(&header[0]).text, "只有首页页眉");
        assert!(parsed.footer.is_none());

        // ③ 没有 sectPr / 没有引用 → None
        let none = parse(&docx_with(r#"<w:p><w:r><w:t>无页眉</w:t></w:r></w:p>"#, &[]));
        assert!(none.header.is_none() && none.footer.is_none());

        // ④ 畸形 r:id、缺部件、rels 指向不存在的文件 → None（不 panic）
        for (rels_xml, sect) in [
            // r:id 指向 rels 里没有的 id
            (rels.clone(), r#"<w:sectPr><w:headerReference w:type="default" r:id="rId404"/></w:sectPr>"#),
            // rels 指向的部件在包里不存在
            (rels.clone(), r#"<w:sectPr><w:headerReference w:type="default" r:id="rId10"/><w:footerReference w:type="default" r:id="rId10"/></w:sectPr>"#),
            // 完全没有 r:id
            (rels.clone(), r#"<w:sectPr><w:headerReference w:type="default"/></w:sectPr>"#),
        ] {
            let body = format!("<w:p/>{}", sect);
            let extras: Vec<(&str, &str)> = vec![("word/_rels/document.xml.rels", &rels_xml)];
            let parsed = parse(&docx_with(&body, &extras));
            // 前两种：部件根本不存在 → None；第三种：没有 id → None
            assert!(
                parsed.header.is_none() || !parsed.header.as_ref().unwrap().is_empty(),
                "要么 None 要么有内容"
            );
        }

        // 页眉部件本身是垃圾 XML → 也不 panic（给空块列表或 None）
        let parsed = parse(&docx_with(
            r#"<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>"#,
            &[
                ("word/_rels/document.xml.rels", &rels),
                ("word/header1.xml", "<<<这不是 XML"),
            ],
        ));
        assert!(parsed.header.map(|blocks| blocks.is_empty()).unwrap_or(true));

        // ⑤ 部件存在但里面什么都没有（全是空段落）→ None，前端不用为它留白
        let empty_footer = format!(r#"<?xml version="1.0"?><w:ftr {NS}><w:p/><w:p/><w:p/></w:ftr>"#);
        let parsed = parse(&docx_with(
            r#"<w:p/><w:sectPr><w:footerReference w:type="default" r:id="rId9"/></w:sectPr>"#,
            &[
                ("word/_rels/document.xml.rels", &rels),
                ("word/footer1.xml", &empty_footer),
            ],
        ));
        assert!(parsed.footer.is_none(), "空页脚当作没有");

        // ⑥ 但"有东西"的页脚要保留（哪怕只有一张图）
        let image_footer = format!(
            r#"<?xml version="1.0"?><w:ftr {NS}><w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic>
              </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:ftr>"#
        );
        let rels_with_image = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
              <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/footer.png"/>
            </Relationships>"#
        );
        let parsed = parse(&docx_with(
            r#"<w:p/><w:sectPr><w:footerReference w:type="default" r:id="rId9"/></w:sectPr>"#,
            &[
                ("word/_rels/document.xml.rels", &rels_with_image),
                ("word/footer1.xml", &image_footer),
            ],
        ));
        let footer = parsed.footer.expect("只有一张图的页脚也算有内容");
        assert!(matches!(footer[0], Block::Image(_)), "页脚里的图片：{footer:?}");
    }

    /* ------------------------------ 实时页码域 ------------------------------ */

    /// **实时页码域**：`PAGE` / `NUMPAGES` 的缓存结果要打上标记，前端用实际页序替换；
    /// 域指令（`w:instrText`）永远不当正文输出。
    #[test]
    fn marks_page_and_numpages_field_runs() {
        let body = r#"
            <w:p>
              <w:r><w:t>第 </w:t></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve">PAGE  </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>2</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:t> 页 共 </w:t></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText>NUMPAGES \* Arabic</w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>30</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:t> 页</w:t></w:r>
            </w:p>
            <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText>DATE \@ "yyyy-MM-dd"</w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>2026-10-21</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:p>
            <w:p><w:r><w:t>普通正文</w:t></w:r></w:p>
            <w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple></w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[]));

        // ① PAGE 域：缓存结果 run 标记为 PAGE，文本照旧是缓存值
        let paragraph = paragraph_of(&blocks[0]);
        assert_eq!(paragraph.text, "第 2 页 共 30 页", "域缓存文本要照常显示");
        let summary: Vec<(String, Option<FieldKind>)> = paragraph
            .runs
            .iter()
            .map(|run| (run.text.clone(), run.field))
            .collect();
        assert_eq!(
            summary,
            vec![
                ("第 ".to_string(), None),
                ("2".to_string(), Some(FieldKind::Page)),
                (" 页 共 ".to_string(), None),
                ("30".to_string(), Some(FieldKind::NumPages)),
                (" 页".to_string(), None),
            ],
            "只有缓存结果的 run 带标记，begin/separate/end 本身不产出 run"
        );
        assert!(
            paragraph.runs.iter().all(|run| !run.text.contains("PAGE")
                && !run.text.contains("NUMPAGES")
                && !run.text.contains("MERGEFORMAT")),
            "域指令绝不能当正文输出"
        );

        // ⑤ 普通域（DATE）：行为完全不变 —— 缓存文本照旧显示，field 为 null
        let date = paragraph_of(&blocks[1]);
        assert_eq!(date.text, "2026-10-21");
        assert_eq!(date.runs.len(), 1);
        assert_eq!(date.runs[0].field, None);

        // ⑥ 普通正文：全部 null
        assert!(paragraph_of(&blocks[2]).runs.iter().all(|run| run.field.is_none()));

        // 附：属性式域 `w:fldSimple`（WPS 常用）里的缓存结果也要标记
        let simple = paragraph_of(&blocks[3]);
        assert_eq!(simple.text, "7");
        assert_eq!(simple.runs[0].field, Some(FieldKind::Page));

        // 前端契约：字段名与取值
        let json = serde_json::to_value(&blocks[0]).expect("应能序列化");
        assert_eq!(json["runs"][1]["field"], "PAGE");
        assert_eq!(json["runs"][3]["field"], "NUMPAGES");
        assert!(json["runs"][0]["field"].is_null());
    }

    /// 嵌套域按**深度计数**处理（`IF` 里套 `PAGE`），只有 `begin` 的域不产出指令文本
    #[test]
    fn nested_and_unseparated_fields_are_handled() {
        // ③ 嵌套：外层的 IF 域里套一个 PAGE 域
        let body = r#"
            <w:p>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText>IF </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText>PAGE</w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>2</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:instrText> = 1 "" "x"</w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>IF 的结果</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:p>
            <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText>PAGE</w:instrText></w:r>
              <w:r><w:t>未算过的域</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>结束</w:t></w:r></w:p>"#;
        let blocks = blocks_of(&docx_with(body, &[]));

        let nested = paragraph_of(&blocks[0]);
        let summary: Vec<(String, Option<FieldKind>)> = nested
            .runs
            .iter()
            .map(|run| (run.text.clone(), run.field))
            .collect();
        assert_eq!(
            summary,
            vec![
                ("2".to_string(), Some(FieldKind::Page)),
                ("IF 的结果".to_string(), None),
            ],
            "嵌套时按最内层已算过的域判断：PAGE 的缓存结果被标记，IF 自己的结果不标记"
        );

        // ④ 只有 begin 没有 separate：不当正文输出指令，也没有任何标记
        let open = paragraph_of(&blocks[1]);
        assert_eq!(open.text, "未算过的域", "指令文本不能漏进正文");
        assert!(open.runs.iter().all(|run| run.field.is_none()));
        // 段落级状态不会漏到下一段（begin 没有 end 也不影响后面的段落）
        assert!(paragraph_of(&blocks[2]).runs.iter().all(|run| run.field.is_none()));
    }

    /// **缓存结果为空的域**（新建 / 从未打印过的文档）：也要产出带标记的空 run，
    /// 否则渲染器没有可替换的对象，页脚页码会是一片空白。
    #[test]
    fn empty_page_fields_still_produce_a_marked_run() {
        let field = |instruction: &str| {
            format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>
                   <w:r><w:instrText>{instruction}</w:instrText></w:r>
                   <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                   <w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            )
        };
        let body = format!(
            r#"{page}{numpages}{both_in_one}{cached}{date_empty}{simple_empty}"#,
            // ① PAGE 域：没有缓存文本 → 一个空 run 带 PAGE 标记
            page = format!("<w:p>{}</w:p>", field("PAGE")),
            // ③ NUMPAGES 同理
            numpages = format!("<w:p>{}</w:p>", field("NUMPAGES \\* Arabic")),
            // 同一段落两个空域 → 两个空 run，按出现顺序
            both_in_one = format!(
                "<w:p><w:r><w:t>第 </w:t></w:r>{}<w:r><w:t> / </w:t></w:r>{}</w:p>",
                field("PAGE"),
                field("NUMPAGES")
            ),
            // ② 有缓存文本 → run 数量不变（不额外补空 run）
            cached = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText>PAGE</w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>7</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
            // 空结果的**其它**域（DATE）不补空 run（只认 PAGE / NUMPAGES）
            date_empty = format!("<w:p>{}</w:p>", field("DATE \\@ \"yyyy-MM-dd\"")),
            // 属性式域（WPS 常用）没算过时同样补一个空 run
            simple_empty = r#"<w:p><w:fldSimple w:instr=" PAGE "/></w:p>"#,
        );
        let blocks = blocks_of(&docx_with(&body, &[]));

        // ① 空 PAGE 域：1 个 run，文本空、带标记
        let page = paragraph_of(&blocks[0]);
        assert_eq!(page.runs.len(), 1, "空结果的域要产出一个空 run：{:?}", page.runs);
        assert_eq!(page.runs[0].text, "");
        assert_eq!(page.runs[0].field, Some(FieldKind::Page));
        // ④ 不让纯文本凭空多出字符
        assert_eq!(page.text, "");
        let json = serde_json::to_value(&blocks[0]).expect("应能序列化");
        assert_eq!(json["runs"][0]["text"], "");
        assert_eq!(json["runs"][0]["field"], "PAGE");
        assert_eq!(json["text"], "");

        // ③ 空 NUMPAGES 同理
        let numpages = paragraph_of(&blocks[1]);
        assert_eq!(numpages.runs.len(), 1);
        assert_eq!(numpages.runs[0].field, Some(FieldKind::NumPages));
        assert_eq!(numpages.text, "");

        // 两个空域在同一段：按出现顺序产出两个空 run，夹在文字之间
        let both = paragraph_of(&blocks[2]);
        let summary: Vec<(String, Option<FieldKind>)> = both
            .runs
            .iter()
            .map(|run| (run.text.clone(), run.field))
            .collect();
        assert_eq!(
            summary,
            vec![
                ("第 ".to_string(), None),
                (String::new(), Some(FieldKind::Page)),
                (" / ".to_string(), None),
                (String::new(), Some(FieldKind::NumPages)),
            ]
        );
        assert_eq!(both.text, "第  / ", "空 run 不改变段落纯文本");

        // ② 有缓存文本时**不多补**空 run（还是那一个缓存 run）
        let cached = paragraph_of(&blocks[3]);
        assert_eq!(cached.runs.len(), 1, "有缓存文本就不该再有空 run");
        assert_eq!((cached.runs[0].text.as_str(), cached.runs[0].field), ("7", Some(FieldKind::Page)));

        // 空结果的 DATE 域：既没有缓存文本、也不补空 run（只认 PAGE / NUMPAGES）
        let date = paragraph_of(&blocks[4]);
        assert!(date.runs.is_empty(), "其它域不补空 run：{:?}", date.runs);
        assert_eq!(date.text, "");

        // 属性式域（`w:fldSimple`）没算过时也一样补空 run
        let simple = paragraph_of(&blocks[5]);
        assert_eq!(simple.runs.len(), 1);
        assert_eq!(simple.runs[0].field, Some(FieldKind::Page));
        assert_eq!(simple.text, "");
    }

    /* ------------------------------ 真实语料库（存在才跑） ------------------------------ */

    /// 拿真实语料库（默认是 Z 盘上的每周汇报目录）跑一遍块模型：
    /// 每个文件都必须能解析成块、每个块都必须能序列化给前端。
    /// 目录不存在时跳过（别的机器 / CI 上没有这个目录）。
    ///
    /// 注意：Word 打开文档时会生成 `~$xxx.docx` **临时锁文件**（不是 zip，也不是文档），
    /// 这里直接跳过 —— 它们不是"解析失败"，是根本不该被当成文档。
    #[test]
    fn parses_real_docx_corpus_without_panicking() {
        fn collect(dir: &Path, out: &mut Vec<std::path::PathBuf>) {
            let Ok(entries) = std::fs::read_dir(dir) else {
                return;
            };
            for entry in entries.flatten() {
                let path = entry.path();
                if path.is_dir() {
                    collect(&path, out);
                } else if path
                    .extension()
                    .is_some_and(|ext| ext.eq_ignore_ascii_case("docx"))
                    && !path
                        .file_name()
                        .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
                {
                    out.push(path);
                }
            }
            out.sort();
        }
        // 语料库目录：默认扫两个（每周汇报 + 项目总结），可用 MASTEREDIT_DOCX_CORPUS 覆盖
        let dirs: Vec<std::path::PathBuf> = match std::env::var("MASTEREDIT_DOCX_CORPUS") {
            Ok(value) => vec![std::path::PathBuf::from(value)],
            Err(_) => vec![
                std::path::PathBuf::from(r"Z:\D\mywork\05_会议汇报\每周汇报\2026"),
                std::path::PathBuf::from(r"Z:\D\mywork\08_项目总结\AI加持下的家庭能源管理系统"),
            ],
        };
        let mut files = Vec::new();
        for dir in &dirs {
            if !dir.exists() {
                eprintln!("跳过：语料库目录不存在 {}", dir.display());
                continue;
            }
            collect(dir, &mut files);
        }
        if files.is_empty() {
            eprintln!("跳过：目录里没有 .docx");
            return;
        }

        let mut total_blocks = 0usize;
        let mut total_tables = 0usize;
        let mut total_images = 0usize;
        let mut total_numbered = 0usize;
        let mut total_unsupported = 0usize;
        let mut merged_cells = 0usize;
        let mut nested_tables = 0usize;
        let mut total_shapes = 0usize;
        let mut total_bordered = 0usize;
        let mut total_text_boxes = 0usize;
        let mut total_page_fields = 0usize;
        let mut total_numpages_fields = 0usize;
        let mut docs_with_header = 0usize;
        let mut docs_with_footer = 0usize;
        let mut labels: HashMap<String, usize> = HashMap::new();
        let mut paper_sizes: HashMap<String, usize> = HashMap::new();

        for path in &files {
            let label = path
                .file_name()
                .map(|name| name.to_string_lossy().to_string())
                .unwrap_or_default();
            let text = path.to_string_lossy().into_owned();
            let parsed = load_document(&text)
                .unwrap_or_else(|error| panic!("{label}: 块模型解析失败：{error}"));

            // 每个块都必须能序列化给前端（前端接的就是这个形状）
            for block in &parsed.blocks {
                serde_json::to_value(block)
                    .unwrap_or_else(|error| panic!("{label}: 块无法序列化：{error}"));
            }
            // 顶层窗口也必须能取（夹取逻辑不 panic）
            let page = document_blocks(text.clone(), 0, MAX_BLOCK_WINDOW)
                .unwrap_or_else(|error| panic!("{label}: document_blocks 失败：{error}"));
            assert_eq!(page.total, parsed.blocks.len(), "{label}: 窗口总数应与块数一致");
            // 页面几何：窗口返回里必须带上（没有 sectPr 的文档是 None，也是合法结果）
            assert_eq!(
                page.page.is_some(),
                parsed.page.is_some(),
                "{label}: 窗口里的 page 应与解析结果一致"
            );
            if let Some(geometry) = parsed.page.as_ref() {
                *paper_sizes
                    .entry(format!(
                        "{:.0}×{:.0}pt{} · 边距 {:.0}/{:.0}/{:.0}/{:.0}",
                        geometry.width_pt,
                        geometry.height_pt,
                        if geometry.landscape { " 横" } else { " 竖" },
                        geometry.margin_top_pt,
                        geometry.margin_right_pt,
                        geometry.margin_bottom_pt,
                        geometry.margin_left_pt,
                    ))
                    .or_insert(0) += 1;
            }
            // 查找也不该 panic（顺带证明索引可用）
            let _ = document_find(text, "的".to_string(), false)
                .unwrap_or_else(|error| panic!("{label}: document_find 失败：{error}"));

            /// 递归统计：表格（含嵌套）、形状、带段落边框的段落、合并单元格
            #[derive(Default)]
            struct Counts {
                tables: usize,
                nested: usize,
                merged: usize,
                shapes: usize,
                bordered: usize,
                text_boxes: usize,
                page_fields: usize,
                numpages_fields: usize,
            }
            fn count_blocks(blocks: &[Block], counts: &mut Counts) {
                for block in blocks {
                    match block {
                        Block::Shape(_) => counts.shapes += 1,
                        Block::TextBox(text_box) => {
                            counts.text_boxes += 1;
                            count_blocks(&text_box.blocks, counts);
                        }
                        Block::Paragraph(paragraph) => {
                            if paragraph.borders.is_some() {
                                counts.bordered += 1;
                            }
                            for run in &paragraph.runs {
                                match run.field {
                                    Some(FieldKind::Page) => counts.page_fields += 1,
                                    Some(FieldKind::NumPages) => counts.numpages_fields += 1,
                                    None => {}
                                }
                            }
                        }
                        Block::Table(table) => {
                            counts.tables += 1;
                            for row in &table.rows {
                                for cell in &row.cells {
                                    if cell.grid_span > 1 || cell.v_merge != VMerge::None {
                                        counts.merged += 1;
                                    }
                                    let before = counts.tables;
                                    count_blocks(&cell.blocks, counts);
                                    counts.nested += counts.tables - before;
                                }
                            }
                        }
                        Block::Image(_) | Block::PageBreak | Block::Unsupported { .. } => {}
                    }
                }
            }
            let mut counts = Counts::default();
            count_blocks(&parsed.blocks, &mut counts);
            // 页眉页脚是独立"故事"，域的标记也要一起数
            if let Some(header) = &parsed.header {
                count_blocks(header, &mut counts);
            }
            if let Some(footer) = &parsed.footer {
                count_blocks(footer, &mut counts);
            }
            let (tables, nested) = (counts.tables, counts.nested);
            merged_cells += counts.merged;
            total_shapes += counts.shapes;
            total_bordered += counts.bordered;
            total_text_boxes += counts.text_boxes;
            total_page_fields += counts.page_fields;
            total_numpages_fields += counts.numpages_fields;
            if parsed.header.is_some() {
                docs_with_header += 1;
            }
            if parsed.footer.is_some() {
                docs_with_footer += 1;
            }
            total_tables += tables;
            total_blocks += parsed.blocks.len();
            nested_tables += nested;
            total_images += parsed
                .blocks
                .iter()
                .filter(|block| matches!(block, Block::Image(_)))
                .count();
            total_numbered += parsed
                .blocks
                .iter()
                .filter(|block| matches!(block, Block::Paragraph(p) if p.list.is_some()))
                .count();
            for block in &parsed.blocks {
                if let Block::Unsupported { label, .. } = block {
                    total_unsupported += 1;
                    *labels.entry(label.clone()).or_insert(0) += 1;
                }
            }
            if counts.shapes > 0 || counts.bordered > 0 || counts.text_boxes > 0 {
                eprintln!(
                    "         └ 形状 {} 个 · 带边框段落 {} 个 · 文本框 {} 个",
                    counts.shapes, counts.bordered, counts.text_boxes
                );
            }
            if counts.page_fields > 0 || counts.numpages_fields > 0 {
                eprintln!(
                    "         └ 域标记：PAGE {} 个 run · NUMPAGES {} 个 run",
                    counts.page_fields, counts.numpages_fields
                );
            }
            let header_text = parsed
                .header
                .as_ref()
                .map(|blocks| blocks.iter().map(block_plain_text).collect::<Vec<_>>().join(" / "))
                .unwrap_or_default();
            let footer_text = parsed
                .footer
                .as_ref()
                .map(|blocks| blocks.iter().map(block_plain_text).collect::<Vec<_>>().join(" / "))
                .unwrap_or_default();
            if !header_text.trim().is_empty() || !footer_text.trim().is_empty() {
                eprintln!(
                    "         └ 页眉「{}」· 页脚「{}」",
                    header_text.trim(),
                    footer_text.trim()
                );
            }
            eprintln!(
                "  [OK]   {label:<30} 块 {:>4} · 段落统计 {:>4} · 表 {:>2} · 图 {:>3} · 形状 {:>3} · 边框段 {:>3} · 文本框 {:>3}",
                parsed.blocks.len(),
                parsed.paragraph_count,
                tables,
                parsed
                    .blocks
                    .iter()
                    .filter(|block| matches!(block, Block::Image(_)))
                    .count(),
                counts.shapes,
                counts.bordered,
                counts.text_boxes
            );
        }

        let mut sorted: Vec<(String, usize)> = labels.into_iter().collect();
        sorted.sort_by(|a, b| b.1.cmp(&a.1));
        eprintln!(
            "\n语料库块模型统计：{} 个文件；顶层块 {}、表格 {}（嵌套 {}）、图片 {}、编号段落 {}、合并单元格 {}、占位块 {}、形状 {}、带边框段落 {}、文本框 {}；有页眉的 {} 个、有页脚的 {} 个；域标记 PAGE {} 个 run / NUMPAGES {} 个 run",
            files.len(),
            total_blocks,
            total_tables,
            nested_tables,
            total_images,
            total_numbered,
            merged_cells,
            total_unsupported,
            total_shapes,
            total_bordered,
            total_text_boxes,
            docs_with_header,
            docs_with_footer,
            total_page_fields,
            total_numpages_fields
        );
        for (label, count) in &sorted {
            eprintln!("  占位类型 {count:>3} × {label}");
        }
        let mut papers: Vec<(String, usize)> = paper_sizes.into_iter().collect();
        papers.sort_by(|a, b| b.1.cmp(&a.1));
        for (size, count) in &papers {
            eprintln!("  纸张 {count:>3} 个 × {size}");
        }
        assert!(!files.is_empty());
    }

    /* ------------------------------ 真实样本（存在才跑） ------------------------------ */

    /// **真实的企业标准封面**（用户实测反馈里"封面差异太大"的那份）：整页由 4 个文本框 + 2 条线组成。
    /// 断言封面文字真的进了文本框块、标题可被查找命中、页眉拿到标准号。文件不存在时静默跳过。
    #[test]
    fn parses_enterprise_standard_cover_if_present() {
        let path = std::env::var("MASTEREDIT_DOCX_ENTERPRISE").unwrap_or_else(|_| {
            r"Z:\D\mywork\08_项目总结\AI加持下的家庭能源管理系统\AI加持下的家庭能源管理系统企业标准.docx"
                .to_string()
        });
        if !Path::new(&path).exists() {
            eprintln!("跳过：样本不存在 {path}");
            return;
        }
        let parsed = load_document(&path).expect("企业标准应能解析");

        // 封面 = 4 个文本框（标准号 / 发布单位 / 发布实施日期 / 标题）+ 2 条分隔线
        let text_boxes: Vec<&TextBoxBlock> = parsed
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::TextBox(text_box) => Some(text_box),
                _ => None,
            })
            .collect();
        assert_eq!(text_boxes.len(), 4, "封面 4 个文本框都要结构化");
        let cover_text: String = text_boxes
            .iter()
            .flat_map(|text_box| text_box.blocks.iter().map(block_plain_text))
            .collect::<Vec<_>>()
            .join("\n");
        for expected in [
            "Q/NDB",
            "AI加持下的家庭能源管理系统",
            "江苏林洋能源股份有限公司",
            "2026-10-21",
        ] {
            assert!(cover_text.contains(expected), "封面文字缺 {expected}：{cover_text}");
        }
        assert!(text_boxes.iter().all(|text_box| text_box.width_pt > 10.0));
        assert!(text_boxes.iter().any(|text_box| text_box.height_pt > 300.0), "标题框很高");
        assert_eq!(text_boxes[0].fill_color.as_deref(), Some("FFFFFF"));

        let shapes: Vec<&ShapeBlock> = parsed
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Shape(shape) => Some(shape),
                _ => None,
            })
            .collect();
        assert_eq!(shapes.len(), 2, "封面两条分隔线");
        assert!(shapes
            .iter()
            .all(|shape| shape.shape == "line" && shape.width_pt > 400.0));

        // 页眉：标准号在**第 1 节**的 default 页眉里，最后一节没有页眉引用 —— 必须往前找到
        // （顺带证明多节文档的页眉不会被"最后一节只有页码"吃掉）
        let header = parsed.header.clone().expect("页眉应能通过多节回退找到");
        let header_text = header.iter().map(block_plain_text).collect::<Vec<_>>().join(" ");
        assert!(header_text.contains("Q/320681NDBXX—2026"), "页眉文本：{header_text}");
        // 页脚是 PAGE 域（缓存结果是数字）—— 标记要打到缓存结果的 run 上，前端好换成实时页码
        let footer = parsed.footer.clone().expect("页脚");
        let footer_text = footer.iter().map(block_plain_text).collect::<Vec<_>>().join(" ");
        assert!(!footer_text.trim().is_empty(), "页脚不该为空：{footer_text:?}");
        let field_runs: Vec<(String, Option<FieldKind>)> = footer
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(&paragraph.runs),
                _ => None,
            })
            .flatten()
            .map(|run| (run.text.clone(), run.field))
            .collect();
        assert!(
            field_runs.iter().any(|(text, field)| *field == Some(FieldKind::Page)
                && text.trim().chars().all(|ch| ch.is_ascii_digit())),
            "页脚的 PAGE 域缓存结果要带标记：{field_runs:?}"
        );

        // 封面文字要能被查找命中（文本框内容算正文）
        let hits =
            document_find(path.clone(), "家庭能源管理系统".to_string(), false).expect("查找应成功");
        assert!(
            hits.iter().any(|hit| hit.block < 6),
            "封面标题应能被搜到（命中块 {:?}）",
            hits.iter().map(|hit| hit.block).collect::<Vec<_>>()
        );
    }

    /// 真实样本：亿赛通加密的周报。断言块模型与 `document_info` 的统计对得上，
    /// 并且样式链、编号、表格、图片都真的解析出来了。样本不存在时静默跳过。
    #[test]
    fn parses_real_encrypted_sample_if_present() {
        let sample = std::env::var("MASTEREDIT_DOCX_SAMPLE")
            .unwrap_or_else(|_| r"C:\Users\master\Documents\7.20-24周报.docx".to_string());
        let path = Path::new(&sample);
        if !path.exists() {
            eprintln!("跳过：样本不存在 {sample}");
            return;
        }
        let path_string = sample.clone();
        let info = document_info(path_string.clone()).expect("document_info 应成功");
        let parsed = parse_document(
            &read_document_bytes(path).expect("应能读取样本").0,
            true,
        )
        .expect("块模型解析应成功");

        // 1) 段落数与既有统计一致（`w:p` 总数，含表格单元格内的）
        assert_eq!(
            parsed.paragraph_count, info.paragraphs,
            "块模型数出来的段落数应与 document_info 一致"
        );
        assert_eq!(info.paragraphs, 319);

        // 2) 表格
        let tables: Vec<&TableBlock> = parsed
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Table(table) => Some(table),
                _ => None,
            })
            .collect();
        assert_eq!(tables.len(), 2, "样本里应有 2 个表格");
        assert_eq!(tables.len(), info.tables);
        assert_eq!(tables[0].rows.len(), 20);
        assert_eq!(tables[1].rows.len(), 7);
        assert_eq!(tables[0].columns.len(), 8);
        assert_eq!(tables[1].columns.len(), 7);
        assert!(tables[0].rows[0].header);
        assert_eq!(tables[0].rows[0].cells[1].text, "论文名称");
        assert_eq!(
            tables[0].rows[0].cells[1].blocks.iter().map(block_plain_text).collect::<Vec<_>>(),
            vec!["论文名称".to_string()]
        );

        // 3) 图片：2 张，路径来自 rels，尺寸由 EMU 换算
        let images: Vec<&ImageBlock> = parsed
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Image(image) => Some(image),
                _ => None,
            })
            .collect();
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].media, "word/media/image1.png");
        assert_eq!(images[1].media, "word/media/image2.png");
        assert_eq!(images[0].name.as_deref(), Some("图片 1"));
        assert!((images[0].width_px - 378.15).abs() < 1.0, "{}", images[0].width_px);
        assert!((images[0].height_px - 242.42).abs() < 1.0, "{}", images[0].height_px);

        // 3.1) 形状与段落边框：这份周报里应该都没有（线条/方框是合同类文档才有的）
        let shapes = parsed
            .blocks
            .iter()
            .filter(|block| matches!(block, Block::Shape(_)))
            .count();
        let bordered = parsed
            .blocks
            .iter()
            .filter(|block| matches!(block, Block::Paragraph(p) if p.borders.is_some()))
            .count();
        eprintln!("样本形状块 {shapes} 个、带段落边框的段落 {bordered} 个");
        assert_eq!(shapes, 0, "周报里没有画线/方框");
        assert_eq!(bordered, 0, "周报里没有段落边框");

        // 4) 样式链：标题 4（styleId=4）应拿到「宋体 / 12pt / 加粗」
        let heading = parsed
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) if paragraph.style_id.as_deref() == Some("4") => {
                    Some(paragraph)
                }
                _ => None,
            })
            .next()
            .expect("样本里应有标题 4 的段落");
        assert_eq!(heading.style.as_deref(), Some("heading 4"));
        assert_eq!(heading.outline_level, Some(3), "标题 4 的大纲级别是 3");
        let run = heading.runs.first().expect("标题段应有文字");
        assert_eq!(run.font_east_asia.as_deref(), Some("宋体"), "样式里的 eastAsia 字体");
        assert_eq!(run.size_pt, Some(12.0), "样式里的 sz=24 半磅");
        assert!(run.bold);

        // 5) 主题字体：document.xml / styles.xml 里**一次都没有"等线"**（只有 theme1.xml 有），
        //    所以能把 run 的字体解析成"等线"就证明主题字体链真的打通了
        //    （WPS 习惯把算好的字体名直接写进 rFonts，那些 run 拿到的是"宋体"）。
        let all_runs = || {
            parsed
                .blocks
                .iter()
                .filter_map(|block| match block {
                    Block::Paragraph(paragraph) => Some(paragraph.runs.iter()),
                    _ => None,
                })
                .flatten()
        };
        let from_theme = all_runs()
            .find(|run| run.font.as_deref() == Some("等线"))
            .expect("应有 run 的西文字体来自主题（docDefaults 只写了 asciiTheme）");
        assert_eq!(from_theme.font_east_asia.as_deref(), Some("等线"), "a:ea 为空 → Hans 字体");
        assert_eq!(from_theme.size_pt, Some(10.5), "docDefaults 的 sz=21 半磅 → 10.5pt");

        let explicit = all_runs()
            .find(|run| run.font.as_deref() == Some("宋体"))
            .expect("样本里大量 run 显式写了字体名");
        assert_eq!(explicit.font_east_asia.as_deref(), Some("宋体"));
        // 这些 run 通常自带 sz（标题/正文各有各的），所以这里只断言字体来源，不断言字号
        assert!(explicit.size_pt.is_some(), "显式写了字体的 run 一般也写了字号");

        // 5.1) 段落的直接格式：样本里有 37 处 line=400 exact → 20pt 固定行距
        let exact = parsed
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => paragraph.line_spacing.as_ref(),
                _ => None,
            })
            .find(|line| line.kind == "exact")
            .expect("样本里有固定行距");
        assert!((exact.value - 20.0).abs() < 0.01, "400 twip = 20pt");

        // 6) 编号：16 个编号段落，前缀按各级格式算出（含中文数字与全角括号）
        let lists: Vec<ListInfo> = parsed
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => paragraph.list.clone(),
                _ => None,
            })
            .collect();
        let prefixes: Vec<String> = lists.iter().map(|list| list.prefix.clone()).collect();
        eprintln!("样本编号前缀：{prefixes:?}");
        assert_eq!(lists.len(), 16, "样本里应有 16 个编号段落");
        assert_eq!(
            prefixes,
            vec![
                "一、", "1、", "2、", "二、", "1、", "2、", "（1）", "（2）", "（3）", "（1）", "（2）",
                "（3）", "（1）", "（2）", "（3）", "（4）"
            ],
            "多级/多套编号的前缀文本"
        );
        assert!(lists.iter().all(|list| list.ordered), "样本里都是有序编号");

        // 7) 分页符：2 个独立分页块；占位块：这个文档不应有
        let page_breaks = parsed
            .blocks
            .iter()
            .filter(|block| matches!(block, Block::PageBreak))
            .count();
        assert_eq!(page_breaks, 2);

        // 7.1) 页面几何：样本是 **A4 横向**（pgSz 16838×11906 twip + orient=landscape，
        //      页边距上下 1800 twip = 90pt、左右 1440 twip = 72pt）
        let geometry = parsed.page.clone().expect("样本里有 w:sectPr");
        assert!(geometry.landscape, "样本是横向");
        assert!((geometry.width_pt - 841.9).abs() < 0.01, "{}", geometry.width_pt);
        assert!((geometry.height_pt - 595.3).abs() < 0.01, "{}", geometry.height_pt);
        assert_eq!(geometry.margin_top_pt, 90.0);
        assert_eq!(geometry.margin_right_pt, 72.0);
        assert_eq!(geometry.margin_bottom_pt, 90.0);
        assert_eq!(geometry.margin_left_pt, 72.0);
        eprintln!("样本页面几何：{geometry:?}");
        let unsupported: Vec<&str> = parsed
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Unsupported { label, .. } => Some(label.as_str()),
                _ => None,
            })
            .collect();
        assert!(unsupported.is_empty(), "这个样本没有不支持的对象：{unsupported:?}");

        // 8) 命令行窗口与查找在真实文档上也能用
        let page = document_blocks(path_string.clone(), 0, 5).expect("窗口取块应成功");
        assert_eq!(page.total, parsed.blocks.len());
        assert_eq!(page.blocks.len(), 5);
        assert!(page.encrypted, "样本是亿赛通加密的");
        let hits = document_find(path_string, "数据集".to_string(), false).expect("查找应成功");
        assert!(!hits.is_empty(), "样本里应有「数据集」的命中");
        assert!(hits[0].block < page.total);

        // 9) 缓存：同一文件再取一次应命中缓存（结果一致）
        let again = document_blocks(sample, 0, 5).expect("再取一次应成功");
        assert_eq!(again.total, page.total);

        let paragraph_blocks = parsed
            .blocks
            .iter()
            .filter(|block| matches!(block, Block::Paragraph(_)))
            .count();
        eprintln!(
            "真实样本：块总数 {}（段落 {} / 表格 {} / 图片 {} / 分页 {} / 占位 {}），编号段落 {}，第一段前缀 {:?}",
            parsed.blocks.len(),
            paragraph_blocks,
            tables.len(),
            images.len(),
            page_breaks,
            unsupported.len(),
            lists.len(),
            prefixes.first()
        );
    }
}
