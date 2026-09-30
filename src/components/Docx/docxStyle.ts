/**
 * DOCX 渲染的「块字段 → CSS」映射与**高度估算**（纯函数，无 React 依赖）。
 *
 * 定位（见 `docs/plan-docx.md`）：结构忠实，不做像素级还原 —— 所以这里只做两件必须做对的事：
 *  1. 把块模型里的排版字段翻译成内联 style（字体/字号/颜色/缩进/间距/行距…）；
 *  2. 给出每块的**估算高度**，供虚拟滚动在「尚未渲染、无法测量」时也能定位。
 *
 * ── 高度估算策略 ──────────────────────────────────────────────────────
 * 段落高度依赖换行，而换行只有浏览器知道。所以走「先估算、渲染后测量修正」：
 *   · 估算：可用宽度 ÷ 文本宽度 → 行数 × 行高 + 段前段后（见 estimateParagraphHeight）；
 *     文本宽度按「中日韩/全角 = 1 em、西文 = 0.52 em」逐字符累加，宁可略大不可略小
 *     （估小了会让后面的块盖上来，估大了只是多一条缝，且会被测量修正）；
 *   · 测量：块渲染出来后读 offsetHeight 回写（见 DocxBlocks 的 measurePass），
 *     以测量值覆盖估算值，并做滚动锚点补偿。
 * 估算与真实渲染共用同一组函数（lineHeightPx / paragraphBoxStyle），保证两者不会各说各话。
 */
import type { CSSProperties } from "react";
import type {
  DocBlock,
  DocBorderSpec,
  DocImage,
  DocLineSpacing,
  DocParagraph,
  DocRun,
  DocShape,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocTextBox,
} from "../../types";
import { buildTableGrid } from "./docxGrid";

/* ============================== 常量 ============================== */

/** pt → px（96dpi：1pt = 4/3 px） */
export const PT_TO_PX = 96 / 72;
/** 文档里没写字号时的兜底：Word 中文默认五号 = 10.5pt */
export const DEFAULT_FONT_PT = 10.5;
/**
 * 单倍行距的行高系数。Word 的单倍行距由字体自身度量决定（中文约 1.3 倍字号），
 * 浏览器 line-height:normal 也差不多，但我们要**可计算**的值才能估算高度，故固定为 1.32。
 */
const LINE_FACTOR = 1.32;
/** 上标/下标字号缩放（Word 约 65%~72%） */
const VERT_ALIGN_SCALE = 0.72;
/** 页面宽度（px）：A4 = 210mm ≈ 794px（后端块模型里没有页面尺寸，用固定值） */
export const PAGE_WIDTH = 794;
/** 页面左右留白（px）：≈ Word 默认页边距 */
export const PAGE_PADDING = 48;
/** 列表每一级的附加缩进（px） */
export const LIST_LEVEL_INDENT = 24;
/** 未加载（还没取回来）的块按这个高度占位 */
export const DEFAULT_BLOCK_HEIGHT = 24;
/** 形状没有给线宽时的默认线宽（pt，≈1px） */
export const DEFAULT_LINE_WIDTH_PT = 0.75;
/** 分页提示线的高度（px） */
export const PAGE_BREAK_HEIGHT = 18;
/** 表格 / 图片 / 占位卡片上下各留一点缝，估算与渲染都用同一组常量 */
const TABLE_MARGIN_Y = 6;
const IMAGE_MARGIN_Y = 6;
const UNSUPPORTED_MARGIN_Y = 5;
/**
 * 表格每一行边界在 `border-collapse: collapse` 下也要占约 1px（实测 4 行表格差 ~4px）。
 * 这是**真实浏览器量出来的**系统性偏差，写进估算里可以让首屏更准（不必等测量修正）。
 */
const TABLE_ROW_BORDER = 1;
/** 图片有 alt 时会多渲染一行图注（11px 字 + 3px 间距，实测 ~20px） */
const IMAGE_CAPTION_HEIGHT = 20;
/** 单元格内边距（px）：渲染与估算共用 */
export const CELL_PADDING_X = 6;
export const CELL_PADDING_Y = 3;
/** 行高最小值（px），防止空表格行塌成一条线 */
const MIN_ROW_HEIGHT = 22;

/** 系统字体回退链（docx 里的字体名在本机缺失时用） */
const SYSTEM_FALLBACK_FONTS = [
  "Microsoft YaHei",
  "微软雅黑",
  "PingFang SC",
  "Hiragino Sans GB",
  "Noto Sans CJK SC",
  "Segoe UI",
  "sans-serif",
];

/* ============================== 小工具 ============================== */

export const round2 = (value: number): number => Math.round(value * 100) / 100;

export function ptToPx(pt: number, scale: number): number {
  return pt * PT_TO_PX * scale;
}

/** `RRGGBB` / `#RRGGBB` → `#RRGGBB`；非法值返回 null（避免把脏值塞进 style） */
export function hexColor(value: string | null): string | null {
  if (!value) return null;
  const raw = value.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(raw)) return null;
  if (raw.toLowerCase() === "auto") return null;
  return `#${raw.toUpperCase()}`;
}

/** Word 页面默认底色：**白色**（即使应用整体处于深色主题，纸面默认也是白纸黑字） */
export const DEFAULT_PAGE_BG = "#ffffff";

/** 页面底色预设（白纸 / 暖黄 / 护眼绿 / 浅灰 / 深色） */
export const DOCX_PAGE_BG_PRESETS: ReadonlyArray<{ id: string; label: string; color: string }> = [
  { id: "white", label: "白色", color: "#ffffff" },
  { id: "warm", label: "暖黄", color: "#faf4e8" },
  { id: "green", label: "护眼绿", color: "#cce8cf" },
  { id: "gray", label: "浅灰", color: "#f3f4f6" },
  { id: "dark", label: "深色", color: "#1e222a" },
];

/** 归一化页面底色（`#rrggbb` 小写），非法或空值回退为默认白色 `#ffffff` */
export function normalizePageBg(value: string | null | undefined): string {
  if (!value) return DEFAULT_PAGE_BG;
  const raw = value.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(raw)) return DEFAULT_PAGE_BG;
  return `#${raw.toLowerCase()}`;
}

/**
 * 按纸面底色亮度决定默认前景色（对应 Word `w:color="auto"`）：
 * 浅色纸面（白纸/暖黄/护眼绿/浅灰）统一用深色墨水 `#1f2328`，避免在应用深色主题下出现「白纸白字」；
 * 深色纸面用浅色字 `#e5e7eb`。段落 run 若显式指定了颜色（`run.color`），仍以文档颜色为准。
 */
export function pageFgColor(bgHex: string): string {
  const norm = normalizePageBg(bgHex).slice(1);
  const r = Number.parseInt(norm.slice(0, 2), 16) / 255;
  const g = Number.parseInt(norm.slice(2, 4), 16) / 255;
  const b = Number.parseInt(norm.slice(4, 6), 16) / 255;
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return lum >= 0.45 ? "#1f2328" : "#e5e7eb";
}

/** CSS 字体名：含非 ASCII / 空格的名字要加引号 */
function quoteFont(name: string): string {
  const trimmed = name.trim();
  if (!trimmed) return "";
  if (/^[A-Za-z][A-Za-z0-9-]*$/.test(trimmed)) return trimmed;
  return `"${trimmed.replace(/"/g, "")}"`;
}

/**
 * run 的 font-family 列表：**西文字体在前、中日韩字体在后**。
 *
 * 这正是 Word 的取字规则：数字、字母、**空格**、西文标点按 `w:ascii`（如 Times New Roman），
 * 汉字按 `w:eastAsia`（如 宋体）。CSS 会为每个字符选**第一个含有该字形的字体**，
 * 所以西文字体在前时汉字自然回落到中文字体，两者各得其所。
 *
 * **真实文档踩过的坑**：早先把 `fontEastAsia` 放在前面，于是空格也用了宋体 ——
 * 宋体的空格是 0.5em，Times New Roman 的只有 0.25em。企业标准封面「2026-10-21发布 +
 * 42 个空格 + 2026-10-21实施」这一行因此多出 196px，把最后一个「施」挤到第二行。
 * 数字同理（宋体数字是全角宽度）。
 */
/** 连续多少个空白才算"排版用空白"（Word 里用空格把日期顶到两端的常见写法） */
const SPACER_RUN_MIN = 3;
/** 排版空白段里空格的宽度（em）：西文字体的空格就这么窄 */
const SPACER_EM = 0.3;

/**
 * 把 run 文本按"连续 ≥3 个空白"切成片段：空白片段要单独用**西文字体**渲染。
 *
 * 原因：Word 逐字符取字体 —— 空格是 ASCII，走 `w:ascii`（Times New Roman 0.25em）；
 * 汉字走 `w:eastAsia`（宋体）。我们的 CSS 只能按 run 选字体，若整段用宋体，
 * 空格就是 0.5em。企业标准封面日期行「2026-10-21发布 + 39 空格 + 2026-10-21实施」
 * 因此多出 196px，最后一个「施」被挤到第二行（Word 里是一行）。
 */
export function splitRunSegments(text: string): Array<{ text: string; spacer: boolean }> {
  const out: Array<{ text: string; spacer: boolean }> = [];
  let index = 0;
  while (index < text.length) {
    const spacers = /^[ \t]+/.exec(text.slice(index));
    if (spacers && spacers[0].length >= SPACER_RUN_MIN) {
      out.push({ text: spacers[0], spacer: true });
      index += spacers[0].length;
      continue;
    }
    // 普通片段：一直吃到"下一个 ≥3 空白段"之前
    let end = index + 1;
    while (end < text.length) {
      if (text[end] === " " || text[end] === "\t") {
        const next = /^[ \t]+/.exec(text.slice(end));
        if (next && next[0].length >= SPACER_RUN_MIN) break;
      }
      end += 1;
    }
    out.push({ text: text.slice(index, end), spacer: false });
    index = end;
  }
  return out;
}

/** 空白片段用的字体链：**西文字体在前**（空格按 0.25em 排，与 Word 的 w:ascii 一致） */
export function spacerFontFamilies(run: DocRun): string[] {
  return fontFamiliesOf(run.font, run.fontEastAsia);
}

export function runFontFamilies(run: DocRun): string[] {
  return fontFamiliesOf(run.fontEastAsia, run.font);
}

/** 按给定顺序拼字体链（去重 + 系统字体兜底） */
function fontFamiliesOf(primary: string | null, secondary: string | null): string[] {
  const out: string[] = [];
  const push = (name: string | null) => {
    const quoted = name ? quoteFont(name) : "";
    if (quoted && !out.includes(quoted)) out.push(quoted);
  };
  push(primary);
  push(secondary);
  for (const fallback of SYSTEM_FALLBACK_FONTS) push(fallback);
  return out;
}

/**
 * 段落容器（strut 支柱行盒）该用哪种字体：**跟段落主 run 一致**。
 *
 * strut 决定段落的基线位置，它的字体度量必须与真正渲染这些字形的字体一致，
 * 否则行盒会被撑到超过设定的 line-height（实测差 2~5px）。用主 run 的字体链即可：
 * 中文段落 → 中文字体，纯西文段落（封面日期行）→ 西文字体。
 */
export function paragraphStrutFamilies(paragraph: DocParagraph): string[] {
  const reference = paragraphReferenceRun(paragraph);
  return reference ? runFontFamilies(reference) : [];
}

/**
 * 宽字符（占 1 em）的判定。
 *
 * **真实文档踩过的坑**：中文排版里的「“ ” ‘ ’ … — （） 、。」等标点在字体里是**全角**（1 em），
 * 但它们散落在 Unicode 的「常用标点」区（U+2010–U+205E）而不是 CJK 区。早先只列了 CJK 区，
 * 于是这些标点按 0.52 em 计 —— 一段 36 字的合同条款被算成 545px（实际 576px），刚好跨过版心宽度，
 * 估算 1 行、实际 2 行，**下一块直接压上来 31px**（真实文档 `实验室对接方案.docx` #18 实测）。
 * 所以这里把中文标点区一并算成宽字符；宁可估宽（多一条缝、会被测量修正），也不能估窄（重叠）。
 */
const WIDE_CHAR =
  /[\u00A5\u00B0\u00B1\u00D7\u00F7\u2010-\u205E\u2103\u2109\u2116\u2122\u1100-\u115F\u2E80-\u303E\u3041-\u33FF\u3400-\u4DBF\u4E00-\u9FFF\uA000-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE19\uFE30-\uFE6F\uFF00-\uFF60\uFFE0-\uFFE6\u{1F300}-\u{1FAFF}]/u;

/** 西文平均字宽（em）：0.52 是常见无衬线字体的经验值 */
const LATIN_EM = 0.52;
/** 制表符按 1.6 em 计（Word 里跳到下一个制表位，通常 0.74cm ≈ 2 个汉字宽的一半） */
const TAB_EM = 1.6;

/**
 * 估算一段文本的显示宽度（px）。逐字符累加：宽字符 1em、制表符 1.6em、其余 0.52em。
 * 只用于虚拟滚动的高度估算，不参与真实排版。
 *
 * 两处特例都来自真实文档实测：
 * - **空格按 0.52em 算**（虽然西文字体里只有 0.25em）：估窄会让后面的块压上来，估宽只是多一条缝；
 * - **但"连续 ≥3 个空白"按 0.3em 算**：那是排版用的空白段，渲染时会单独用西文字体（见
 *   `splitRunSegments`），按 0.52em 会明显高估（企业标准封面日期行会多算一行）。
 */
export function estimateTextWidthPx(text: string, fontPx: number): number {
  let em = 0;
  let spacerRun = 0; // 当前连续空白计数
  for (const ch of text) {
    if (ch === " " || ch === "\t") {
      spacerRun += 1;
      em += spacerRun >= SPACER_RUN_MIN ? (ch === "\t" ? TAB_EM : 0.3) : LATIN_EM;
      continue;
    }
    spacerRun = 0;
    if (WIDE_CHAR.test(ch)) em += 1;
    else em += LATIN_EM;
  }
  return em * fontPx;
}

/* ============================== 段落：样式 ============================== */

/**
 * 对齐映射。`both` / `distribute` 都用 `justify`：两端对齐的观感接近，
 * 逐字分配（distribute）浏览器只有 `text-justify: distribute`，兼容性差，不值得为它引入风险。
 */
export function alignToCss(align: string | null): CSSProperties["textAlign"] | undefined {
  switch (align) {
    case "left":
      return "left";
    case "center":
      return "center";
    case "right":
      return "right";
    case "both":
    case "distribute":
      return "justify";
    default:
      return undefined;
  }
}

/** 段落的"主字号"（pt）：取 run 里最大的字号，没有 run 时用默认五号 */
export function paragraphFontPt(paragraph: DocParagraph): number {
  let pt = 0;
  for (const run of paragraph.runs) {
    if (run.sizePt !== null && run.sizePt > pt) pt = run.sizePt;
  }
  return pt > 0 ? pt : DEFAULT_FONT_PT;
}

/** 段落里"最大字号"的那个 run（决定行盒的字体与字号，也就是容器的 strut 该用什么） */
export function paragraphReferenceRun(paragraph: DocParagraph): DocRun | null {
  let best: DocRun | null = null;
  let bestPt = 0;
  for (const run of paragraph.runs) {
    const pt = run.sizePt ?? 0;
    if (best === null || pt > bestPt) {
      best = run;
      bestPt = pt;
    }
  }
  return best;
}

/** 单行的行高（px）：倍数行距按 系数×倍数，固定/最小值行距直接用 pt 折算 */
export function lineHeightPx(paragraph: DocParagraph, scale: number): number {
  const fontPx = ptToPx(paragraphFontPt(paragraph), scale);
  const spacing: DocLineSpacing | null = paragraph.lineSpacing;
  if (spacing && spacing.value > 0) {
    if (spacing.kind === "multiple") {
      return Math.max(fontPx, fontPx * LINE_FACTOR * spacing.value);
    }
    // exact / atLeast 的 value 单位是 pt
    return Math.max(fontPx, ptToPx(spacing.value, scale));
  }
  return fontPx * LINE_FACTOR;
}

/**
 * 列表层级的附加缩进（px）。
 *
 * **注意不要双重缩进**：后端在段落属性层叠里已经把 `numbering.xml` 该级别的
 * `w:lvl/w:pPr/w:ind/@w:left` 折算进了 `indentLeftPt`（见 `office_docx.rs` 的
 * 「直接格式 → 编号级别缩进 → 段落样式 → docDefaults」），所以只有在文档**完全没给**
 * 左缩进（`indentLeftPt === null`）时才用 level 兜底。
 */
export function listIndentPx(paragraph: DocParagraph, scale: number): number {
  if (!paragraph.list) return 0;
  if (paragraph.indentLeftPt !== null) return 0;
  return LIST_LEVEL_INDENT * Math.max(0, paragraph.list.level) * scale;
}

/** 列表标记与正文之间的间隔（px） */
export function listGapPx(paragraph: DocParagraph, scale: number): number {
  const fontPx = ptToPx(paragraphFontPt(paragraph), scale);
  if (!paragraph.list) return 0;
  if (paragraph.list.suffix === "nothing") return 0;
  return fontPx * 0.5;
}

/** 列表前缀本身的宽度（px）：用于估算首行悬挂量 */
export function listPrefixWidthPx(paragraph: DocParagraph, scale: number): number {
  if (!paragraph.list) return 0;
  const fontPx = ptToPx(paragraphFontPt(paragraph), scale);
  return estimateTextWidthPx(paragraph.list.prefix, fontPx) + listGapPx(paragraph, scale);
}

/**
 * 首行缩进（px）：**负数 = 悬挂**。
 * 列表段若文档没给首行缩进，就按「前缀宽度」做默认悬挂 —— 这样正文换行后对齐到
 * paddingLeft，编号/项目符号挂在左边，观感与 Word 一致。
 */
export function firstLineIndentPx(paragraph: DocParagraph, scale: number): number {
  if (paragraph.indentFirstLinePt !== null) return ptToPx(paragraph.indentFirstLinePt, scale);
  if (paragraph.list) return -listPrefixWidthPx(paragraph, scale);
  return 0;
}

/* ------------------------------ 段落边框 ------------------------------ */

/** 边框样式映射：`w:val` → CSS border-style（认不出的按实线处理） */
export function borderStyleToCss(style: string | null | undefined): string {
  switch ((style ?? "").toLowerCase()) {
    case "dashed":
    case "dash":
    case "dashsmallgap":
    case "lgdash":
      return "dashed";
    case "dotted":
    case "dot":
    case "sysdot":
    case "dotdash":
      return "dotted";
    case "double":
      return "double";
    default:
      return "solid";
  }
}

/**
 * 段落每条边的「占位」（px）= 线宽 + 间距（`w:space`）。
 * 渲染时把它加到同侧的 padding 上、边框画在外面 —— 这样**边框在盒模型里的占位
 * 与高度估算完全一致**，不会出现"有边框的段落和下一段重叠"。
 */
export function paragraphBorderInsets(
  paragraph: DocParagraph,
  scale: number,
): { top: number; right: number; bottom: number; left: number } {
  const zero = { top: 0, right: 0, bottom: 0, left: 0 };
  const borders = paragraph.borders;
  if (!borders) return zero;
  const side = (spec: DocBorderSpec | null): number =>
    spec ? ptToPx(Math.max(0, spec.widthPt), scale) + ptToPx(Math.max(0, spec.spacePt ?? 0), scale) : 0;
  return {
    top: side(borders.top),
    right: side(borders.right),
    bottom: side(borders.bottom),
    left: side(borders.left),
  };
}

/** 一条边框 → 内联 border-* 值（颜色缺省用 currentColor：Word 的 auto 就是文字色） */
function borderCssValue(spec: DocBorderSpec, scale: number): string {
  const width = Math.max(0.5, round2(ptToPx(Math.max(0, spec.widthPt), scale)));
  const color = hexColor(spec.color) ?? "currentColor";
  return `${width}px ${borderStyleToCss(spec.style)} ${color}`;
}

/** 段落的盒模型样式（缩进、间距、行距、对齐、换行规则、四边边框） */
export function paragraphBoxStyle(paragraph: DocParagraph, scale: number): CSSProperties {
  const lineHeight = lineHeightPx(paragraph, scale);
  const insets = paragraphBorderInsets(paragraph, scale);
  const reference = paragraphReferenceRun(paragraph);
  const style: CSSProperties = {
    /**
     * **容器的字号与字体必须跟着段落主字号/主字体走**：容器自身的"strut"（支柱行盒）
     * 用的是容器的字体度量。若容器沿用应用根字号的 13px + 界面字体，而 run 是 16~22pt 的
     * 宋体/黑体，两者的半行距不同，行盒会被撑到超过指定的 line-height ——
     * 真实文档实测：22pt 标题设定 93.25px、实际 98.33px（+5px）；16pt 标题 +1.8px。
     * 字号对齐只解决一半，**字体也要对齐**（同一字号下不同字体的 ascent/descent 仍不同）。
     * 只影响 strut 与空行高度，run 自己的字体字号照旧生效，观感不变。
     */
    fontSize: `${round2(ptToPx(paragraphFontPt(paragraph), scale))}px`,
    ...(reference ? { fontFamily: paragraphStrutFamilies(paragraph).join(", ") } : {}),
    // pre-wrap：块模型里的 `\n`（软换行）与 `\t`（制表符）要原样保留
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    overflowWrap: "anywhere",
    lineHeight: `${round2(lineHeight)}px`,
    minHeight: `${round2(lineHeight)}px`,
    // 段前段后 + 边框侧的「线宽 + 间距」都进 padding（绝对定位下 margin 不计入 offsetHeight，
    // 用 padding 才能让"估算高度"与"实测高度"是同一个口径）
    paddingTop: `${round2(ptToPx(paragraph.spaceBeforePt ?? 0, scale) + insets.top)}px`,
    paddingBottom: `${round2(ptToPx(paragraph.spaceAfterPt ?? 0, scale) + insets.bottom)}px`,
    paddingLeft: `${round2(ptToPx(paragraph.indentLeftPt ?? 0, scale) + listIndentPx(paragraph, scale) + insets.left)}px`,
    paddingRight: `${round2(ptToPx(paragraph.indentRightPt ?? 0, scale) + insets.right)}px`,
    textIndent: `${round2(firstLineIndentPx(paragraph, scale))}px`,
  };
  const borders = paragraph.borders;
  if (borders?.top) style.borderTop = borderCssValue(borders.top, scale);
  if (borders?.right) style.borderRight = borderCssValue(borders.right, scale);
  if (borders?.bottom) style.borderBottom = borderCssValue(borders.bottom, scale);
  if (borders?.left) style.borderLeft = borderCssValue(borders.left, scale);
  const align = alignToCss(paragraph.align);
  if (align) style.textAlign = align;
  return style;
}

/**
 * 域代码 run 的**实时文本**：`PAGE` → 当前页码，`NUMPAGES` → 总页数。
 *
 * 文档里带的是上次保存/打印时的缓存值（`"2"`、`"30"`），必须换掉才对得上真实分页。
 * 缓存文本**可能为空串**（新建或从未打印过的文档）—— 这种情况同样要替换，
 * 绝不能因为 `text === ""` 就把 run 跳过（那样页码会整段消失）。
 */
export function resolveRunText(run: DocRun, pageNumber: number, totalPages: number): string {
  if (run.field === "PAGE") return String(Math.max(1, Math.round(pageNumber)));
  if (run.field === "NUMPAGES") return String(Math.max(1, Math.round(totalPages)));
  return run.text;
}

/**
 * 域 run 的**稳定占位宽度**（px）：按缓存文本的宽度给一个 `min-width`，
 * 让替换后的数字不会改变行宽 → 不会触发重新换行 → 也就不会"改分页 → 再测 → 再改"地抖动。
 * 缓存文本为空时按一个数字的宽度兜底。
 */
export function fieldRunMinWidthPx(run: DocRun, scale: number): number {
  const fontPx = ptToPx(run.sizePt ?? DEFAULT_FONT_PT, scale);
  const source = run.text && run.text.trim() ? run.text : "0";
  return round2(estimateTextWidthPx(source, fontPx));
}

/** 单个 run 的内联样式：字体、字号、粗斜体、下划线/删除线、颜色、高亮、上下标 */
export function runInlineStyle(run: DocRun, scale: number): CSSProperties {
  const style: CSSProperties = { fontFamily: runFontFamilies(run).join(", ") };
  const isVertAlign = run.vertAlign === "superscript" || run.vertAlign === "subscript";
  const fontPx = ptToPx(run.sizePt ?? DEFAULT_FONT_PT, scale) * (isVertAlign ? VERT_ALIGN_SCALE : 1);
  style.fontSize = `${round2(fontPx)}px`;
  if (run.bold) style.fontWeight = 700;
  if (run.italic) style.fontStyle = "italic";
  const decorations: string[] = [];
  if (run.underline) decorations.push("underline");
  if (run.strike) decorations.push("line-through");
  if (decorations.length > 0) style.textDecoration = decorations.join(" ");
  const color = hexColor(run.color);
  if (color) style.color = color;
  const highlight = hexColor(run.highlight);
  if (highlight) style.backgroundColor = highlight;
  if (run.vertAlign === "superscript") {
    style.verticalAlign = "super";
    // 上下标不应把行高撑大（Word 也不撑）
    style.lineHeight = 0;
  } else if (run.vertAlign === "subscript") {
    style.verticalAlign = "sub";
    style.lineHeight = 0;
  }
  return style;
}

/** 段落是否要画"分页提示线"（段前分页 / 分节符），以及线右侧的说明文字 */
export function breakHintText(paragraph: DocParagraph): string | null {
  if (paragraph.sectionBreak) {
    const labels: Record<string, string> = {
      nextPage: "下一页",
      continuous: "连续",
      evenPage: "偶数页",
      oddPage: "奇数页",
    };
    return `分节符（${labels[paragraph.sectionBreak] ?? paragraph.sectionBreak}）`;
  }
  if (paragraph.pageBreakBefore || paragraph.pageBreak) return "分页";
  return null;
}

/* ============================== 高度估算 ============================== */

/**
 * 单个字符的估算宽度（px），只看它占几个 em。
 * `spacer` = 这个空格属于"连续 ≥3 个空白"的排版空白段（渲染时会单独用西文字体，只有 0.3em）。
 */
function charWidthPx(char: string, fontPx: number, spacer: boolean): number {
  if (char === "\t") return fontPx * TAB_EM;
  if (spacer) return fontPx * SPACER_EM;
  return WIDE_CHAR.test(char) ? fontPx : fontPx * LATIN_EM;
}

/**
 * 单遍扫描估算**行数**：按字符累加宽度，超出可用宽度就换行，遇到 `\n`（软换行）强制换行。
 *
 * 为什么逐字符扫而不是「总宽 ÷ 可用宽」一次除：段落里各 run 字号不同（标题里混小字很常见），
 * 且软换行会把段落切成语义行；一次除法在混排时误差能达到一倍，直接导致滚动定位跑偏。
 *
 * `firstLineIndentPx` 为正（首行缩进）时**首行可用宽度要减掉它** —— 中文合同条款几乎都是
 * 「首行缩进 2 字符」，不减就会把刚好跨行的段落算少一行。
 */
function countLines(
  runs: Array<{ text: string; fontPx: number }>,
  availPx: number,
  leadingWidth: number,
  firstLineIndentPx = 0,
): number {
  const avail = Math.max(24, availPx);
  const availFirst = Math.max(24, avail - Math.max(0, firstLineIndentPx));
  let lines = 1;
  let width = Math.min(leadingWidth, availFirst); // 列表前缀占首行宽度
  let hasContent = false;
  let spacerRun = 0; // 当前连续空白计数（≥3 个按 0.3em 的排版空白算）
  for (const run of runs) {
    for (const char of run.text) {
      if (char === "\n") {
        lines += 1;
        width = 0;
        spacerRun = 0;
        continue;
      }
      if (char === " " || char === "\t") spacerRun += 1;
      else spacerRun = 0;
      hasContent = true;
      const w = charWidthPx(char, run.fontPx, spacerRun >= SPACER_RUN_MIN);
      if (width > 0 && width + w > (lines === 1 ? availFirst : avail)) {
        lines += 1;
        width = 0;
      }
      width += w;
    }
  }
  return hasContent || lines > 1 ? lines : 1;
}

/** 把段落的 run 摊平成「文本 + 该段字号（px）」，供行数估算逐字符累加 */
function paragraphRunSegments(
  paragraph: DocParagraph,
  scale: number,
): Array<{ text: string; fontPx: number }> {
  const segments = paragraph.runs.map((run) => ({
    text: run.text,
    fontPx: ptToPx(run.sizePt ?? DEFAULT_FONT_PT, scale),
  }));
  // run 全空但 text 有内容（理论不该出现）：退化成用段落主字号估整段文本
  if (segments.every((segment) => !segment.text) && paragraph.text) {
    return [{ text: paragraph.text, fontPx: ptToPx(paragraphFontPt(paragraph), scale) }];
  }
  return segments;
}

/** 估算一个段落块的高度（px）：行数 × 行高 + 段前段后 + 边框占位 + 分页提示线 */
export function estimateParagraphHeight(
  paragraph: DocParagraph,
  scale: number,
  contentWidth: number,
): number {
  const lineHeight = lineHeightPx(paragraph, scale);
  const insets = paragraphBorderInsets(paragraph, scale);
  const paddingLeft =
    ptToPx(paragraph.indentLeftPt ?? 0, scale) + listIndentPx(paragraph, scale) + insets.left;
  const paddingRight = ptToPx(paragraph.indentRightPt ?? 0, scale) + insets.right;
  const avail = contentWidth - paddingLeft - paddingRight;
  const lines = countLines(
    paragraphRunSegments(paragraph, scale),
    avail,
    listPrefixWidthPx(paragraph, scale),
    firstLineIndentPx(paragraph, scale),
  );
  const spacing =
    ptToPx(paragraph.spaceBeforePt ?? 0, scale) +
    ptToPx(paragraph.spaceAfterPt ?? 0, scale) +
    insets.top +
    insets.bottom;
  const hint = breakHintText(paragraph) ? PAGE_BREAK_HEIGHT : 0;
  return round2(lines * lineHeight + spacing + hint);
}

/**
 * 估算形状块的高度（px）= `yPt + heightPt`（相对段落内容区左上角）。
 * 线（`heightPt = 0`）按**线宽**兜底，否则会算成 0 高度、和下一块叠在一起。
 */
export function estimateShapeHeight(shape: DocShape, scale: number): number {
  const offsetY = ptToPx(Math.max(0, shape.yPt), scale);
  const own = ptToPx(Math.max(0, shape.heightPt), scale);
  const line = ptToPx(shape.lineWidthPt !== null && shape.lineWidthPt > 0 ? shape.lineWidthPt : DEFAULT_LINE_WIDTH_PT, scale);
  return round2(offsetY + Math.max(own, line, 1));
}

/** 估算单元格内容高度（px） */
function estimateCellHeight(cell: DocTableCell, scale: number, widthPx: number): number {
  const inner = Math.max(40, widthPx - CELL_PADDING_X * 2);
  return estimateBlocksHeight(cell.blocks, scale, inner) + CELL_PADDING_Y * 2;
}

/** 估算一行表格的高度（px）：行高属性与单元格内容高度取大，并夹到最小值 */
function estimateRowHeight(
  row: DocTableRow,
  scale: number,
  columns: number[],
  grid: ReturnType<typeof buildTableGrid>,
  contentWidth: number,
): number {
  let height = row.heightPx !== null && row.heightPx > 0 ? row.heightPx * scale : 0;
  for (const entry of grid.entries) {
    if (entry.row !== row) continue;
    if (entry.continuation) continue; // 纵向延续的格子不渲染 `<td>`，不参与高度
    let width =
      entry.cell.widthPx !== null && entry.cell.widthPx > 0 ? entry.cell.widthPx * scale : 0;
    if (width <= 0) {
      width = 0;
      for (let c = entry.colStart; c < entry.colStart + entry.colSpan; c += 1) {
        width += (columns[c] ?? 0) * scale;
      }
    }
    if (width <= 0) width = contentWidth / Math.max(1, grid.columnCount);
    height = Math.max(height, estimateCellHeight(entry.cell, scale, width));
  }
  return Math.max(MIN_ROW_HEIGHT, height);
}

/** 估算一个表格块的高度（px） */
export function estimateTableHeight(table: DocTable, scale: number, contentWidth: number): number {
  const grid = buildTableGrid(table);
  let total = TABLE_MARGIN_Y * 2 + table.rows.length * TABLE_ROW_BORDER;
  for (const row of table.rows) {
    total += estimateRowHeight(row, scale, table.columns, grid, contentWidth);
  }
  return round2(total);
}

/** 估算一组块的高度（px）：表格单元格 / 嵌套结构复用 */
export function estimateBlocksHeight(blocks: DocBlock[], scale: number, contentWidth: number): number {
  let total = 0;
  for (const block of blocks) total += estimateBlockHeight(block, scale, contentWidth);
  return total;
}

/**
 * 估算占位卡片高度（px）。支持**合并后的多行说明**：
 * 标题行 + 每一条 detail 各占若干行。
 * 卡片是**我们自己的界面元素**（不是文档内容），字号固定 11/12px，所以不吃缩放倍率 ——
 * 渲染与估算都不乘 scale，两边一致。
 */
export function estimateUnsupportedHeight(details: readonly string[], contentWidth: number): number {
  const avail = Math.max(120, contentWidth - 24);
  let lines = 0;
  for (const detail of details) {
    if (!detail) continue;
    lines += countLines([{ text: detail, fontPx: 12 }], avail, 0);
  }
  return round2(32 + lines * 16 + UNSUPPORTED_MARGIN_Y * 2);
}

/** 图片拿不到尺寸时的兜底高度（px）：**估算与渲染必须用同一个值** */
export const IMAGE_FALLBACK_HEIGHT = 160;

/**
 * 图片的显示盒子（px）——**估算与渲染共用这一个函数**，保证"估算高度 = 渲染高度"这条不变式。
 *
 * 两条修过的坑：
 *  1. 图片比版心宽时，只压宽度不缩高度 → 盒子高度 ≠ 估算高度，且图被压扁。
 *     这里按比例**同时缩宽缩高**（`height = height × 版心宽 / 宽`）。
 *  2. `heightPx` 为 0 / 缺失时，估算是 160、渲染却是 `height: undefined`（图片自然高，可能几百 px）
 *     → 下一块直接压上来。这里两边都用 `IMAGE_FALLBACK_HEIGHT` + `object-fit: contain`。
 */
export function resolveImageBox(
  block: DocImage,
  scale: number,
  contentWidth: number,
): { width: number; height: number } {
  const avail = Math.max(40, contentWidth);
  let width = block.widthPx > 0 ? block.widthPx * scale : avail;
  let height = block.heightPx > 0 ? block.heightPx * scale : IMAGE_FALLBACK_HEIGHT;
  if (width > avail) {
    const factor = avail / width;
    width = avail;
    height *= factor;
  }
  return { width: round2(width), height: round2(Math.max(1, height)) };
}

/** 图片读取超时（ms）：超时走失败态，绝不停在「图片加载中…」（`DocxView` 与 `DocxBlock` 共用） */
export const MEDIA_TIMEOUT_MS = 8000;
/** 图片失败后重试的最小间隔（ms）：避免观测器重建导致请求风暴 */
export const MEDIA_RETRY_MS = 2000;

/** 文本框内边距（px）：渲染与估算共用 */
export const TEXT_BOX_PADDING = 6;
/** 文本框边框的默认线宽（pt） */
export const TEXT_BOX_DEFAULT_BORDER_PT = 0.75;

/**
 * 文本框的盒子（px）——**估算与渲染共用**，保证"估算高度 = 渲染高度"。
 *
 * **宽度口径（修过的坑）**：`maxBoxWidth` 是**纸张可容纳宽度**，不是正文版心宽。
 * Word 里文本框是浮动对象，**允许超出正文版心**（只受纸张边界约束）。早先夹到版心
 * （企业标准：声明 482pt = 642.67px，版心只有 623.67px，差 19px）→ 日期行
 * 「2026-10-21实施」被从中间挤断成两行。文本框左边缘仍对齐版心左边缘，超出部分向右溢出。
 *
 * 高度不变式：`height = max(heightPt, 内部内容高 + 内边距×2)` —— **内容优先、绝不裁切**。
 */
export function resolveTextBoxBox(
  box: DocTextBox,
  scale: number,
  maxBoxWidth: number,
  innerContentHeight: number,
): { width: number; height: number; innerContentWidth: number } {
  const avail = Math.max(40, maxBoxWidth);
  const declaredWidth = ptToPx(Math.max(0, box.widthPt), scale);
  const width = Math.min(avail, declaredWidth > 0 ? declaredWidth : avail);
  const innerContentWidth = Math.max(24, width - TEXT_BOX_PADDING * 2);
  const declaredHeight = ptToPx(Math.max(0, box.heightPt), scale);
  const height = Math.max(declaredHeight, innerContentHeight + TEXT_BOX_PADDING * 2);
  return { width: round2(width), height: round2(height), innerContentWidth: round2(innerContentWidth) };
}

/** 估算文本框块的高度（px）：宽度只影响内部换行，高度由内容决定 */
export function estimateTextBoxHeight(
  box: DocTextBox,
  scale: number,
  maxBoxWidth: number,
): number {
  const avail = Math.max(40, maxBoxWidth);
  const declaredWidth = ptToPx(Math.max(0, box.widthPt), scale);
  const width = Math.min(avail, declaredWidth > 0 ? declaredWidth : avail);
  const innerWidth = Math.max(24, width - TEXT_BOX_PADDING * 2);
  const innerHeight = estimateBlocksHeight(box.blocks, scale, innerWidth);
  return resolveTextBoxBox(box, scale, maxBoxWidth, innerHeight).height;
}

/**
 * 估算单块高度（px）。**虚拟滚动在块尚未渲染时只能靠它定位**，
 * 所以任何分支都必须给出一个正数（宁可偏大）。
 *
 * `maxBlockWidth` 只有"允许超出正文版心的浮动对象"（文本框）用得上；
 * 其余块一律用 `contentWidth`（版心 / 单元格内宽）。
 */
export function estimateBlockHeight(
  block: DocBlock,
  scale: number,
  contentWidth: number,
  maxBlockWidth: number = contentWidth,
): number {
  switch (block.kind) {
    case "paragraph":
      return estimateParagraphHeight(block, scale, contentWidth);
    case "table":
      return estimateTableHeight(block, scale, contentWidth);
    case "image": {
      const box = resolveImageBox(block, scale, contentWidth);
      // 有 alt 时还会多一行图注（真实浏览器实测约 20px）
      const caption = block.alt && block.alt.trim() ? IMAGE_CAPTION_HEIGHT : 0;
      return round2(box.height + IMAGE_MARGIN_Y * 2 + caption);
    }
    case "pageBreak":
      return PAGE_BREAK_HEIGHT;
    case "shape":
      return estimateShapeHeight(block, scale);
    case "textBox":
      return estimateTextBoxHeight(block, scale, Math.max(contentWidth, maxBlockWidth));
    case "unsupported":
      return estimateUnsupportedHeight([block.detail], contentWidth);
    default:
      return DEFAULT_BLOCK_HEIGHT;
  }
}

/** 占位卡片 / 图片 / 表格外层容器的上下留白（渲染与估算共用） */
export const BLOCK_MARGIN_Y = {
  table: TABLE_MARGIN_Y,
  image: IMAGE_MARGIN_Y,
  unsupported: UNSUPPORTED_MARGIN_Y,
} as const;
