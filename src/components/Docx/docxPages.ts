/**
 * 分页排版（纯函数）：把有序块**按累计高度装进一页一页**，并给出两种模式共用的
 * 「块 → 滚动容器内 y」坐标。
 *
 * ── 页几何 ────────────────────────────────────────────────────────────
 * 优先用后端给的 `page`（取自第一个 `w:sectPr`，单位 pt）；字段没落地时按
 * **A4 + 2.54cm 页边距**兜底（Word 中文默认就是这个）。pt → px 用 96dpi：×4/3。
 *
 * ── 分页规则 ──────────────────────────────────────────────────────────
 * 1. **显式分页**：`kind === "pageBreak"` 的块、段落上的 `pageBreakBefore` /
 *    `sectionBreak` 一律强制开新页（即使上一页还很空 —— 这正是 Word 的行为）；
 * 2. **自然分页**：当前页放不下下一块（累计高度 + 块高 > 可用内容高）就换页，
 *    但**一页至少放一块**（避免块高恰好等于页高时出现空页死循环）；
 * 3. **超高块**（比一页还高的表格/图片）：**不裁切**，让它独占一页，并把这页的
 *    卡片按内容撑高（`内容高 + 上下边距`）—— 视觉上就是"这一页长了一点"，
 *    而不是与下一页的卡片重叠。选这个策略是因为另两种做法各有硬伤：
 *    真溢出会压住下一页的纸面；按页高切开要自己断表（我们定位是"结构忠实"，
 *    不做像素级断行/断表）。代价是超高块那一页与 Word 的分页位置不同。
 *
 * 坐标：所有块的 top 都换算成**滚动容器坐标**（页面卡片的 top + 页面上边距 + 页内 y），
 * 因此连续模式与分页模式可以用同一套「跳转 / 锚点补偿 / 高度测量」逻辑。
 */

/** pt → px（96dpi） */
const PT_TO_PX = 96 / 72;

/** A4 兜底（pt）：210mm × 297mm */
const A4_WIDTH_PT = 595.28;
const A4_HEIGHT_PT = 841.89;
/** Word 中文默认页边距 2.54cm = 72pt */
const A4_MARGIN_PT = 72;
/** 页与页之间的空隙（px） */
export const PAGE_GAP = 20;
/** 页脚区域高度（px）：页码占的地方，不参与内容排版 */
export const PAGE_FOOTER_HEIGHT = 26;

export interface PageGeometry {
  /** 纸张宽（px） */
  widthPx: number;
  /** 纸张高（px） */
  heightPx: number;
  marginTopPx: number;
  marginRightPx: number;
  marginBottomPx: number;
  marginLeftPx: number;
  /** 可用内容宽（px）：估算与渲染共用，必须一致 */
  contentWidthPx: number;
  /** 可用内容高（px）：一页能放多少内容 */
  contentHeightPx: number;
}

/** 后端给的页面设置（字段名与 `DocPageSetup` 一致，这里不 import 以保持纯函数模块独立） */
export interface PageSetupInput {
  widthPt?: number | null;
  heightPt?: number | null;
  marginTopPt?: number | null;
  marginRightPt?: number | null;
  marginBottomPt?: number | null;
  marginLeftPt?: number | null;
}

/**
 * 解析页几何：**先信后端，缺哪个补哪个**（不是整组回退）——
 * 这样后端只给出页面尺寸、没给页边距时，也不会把版心算错。
 */
export function resolvePageGeometry(setup: PageSetupInput | null | undefined): PageGeometry {
  const widthPt = positive(setup?.widthPt) ?? A4_WIDTH_PT;
  const heightPt = positive(setup?.heightPt) ?? A4_HEIGHT_PT;
  const marginTopPx = ptToPx(positive(setup?.marginTopPt) ?? A4_MARGIN_PT);
  const marginRightPx = ptToPx(positive(setup?.marginRightPt) ?? A4_MARGIN_PT);
  const marginBottomPx = ptToPx(positive(setup?.marginBottomPt) ?? A4_MARGIN_PT);
  const marginLeftPx = ptToPx(positive(setup?.marginLeftPt) ?? A4_MARGIN_PT);
  const widthPx = round2(widthPt * PT_TO_PX);
  const heightPx = round2(heightPt * PT_TO_PX);
  return {
    widthPx,
    heightPx,
    marginTopPx,
    marginRightPx,
    marginBottomPx,
    marginLeftPx,
    // 版心至少要留一点宽度/高度，避免畸形 sectPr 把内容挤成 0
    contentWidthPx: Math.max(120, round2(widthPx - marginLeftPx - marginRightPx)),
    /**
     * 版心高 = 页面高 - 上下边距（**与 Word 一致，不再额外扣页脚**）。
     * 文档自己的页脚画在**下边距里**，不占版心；我们自己的页码已经移到纸张外。
     */
    contentHeightPx: Math.max(80, round2(heightPx - marginTopPx - marginBottomPx)),
  };
}

function positive(value: number | null | undefined): number | null {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : null;
}

function round2(value: number): number {
  return Math.round(value * 100) / 100;
}

function ptToPx(pt: number): number {
  return round2(pt * PT_TO_PX);
}

/* ============================== 分页结果 ============================== */

export interface DocxPage {
  /** 页序号（0 起） */
  index: number;
  /** 该页第一个块下标 */
  start: number;
  /** 该页最后一个块下标（含） */
  end: number;
  /** 页内内容高度（px，不含页边距） */
  contentHeight: number;
  /** 纸张卡片高度（px，含上下边距；超高块那页会被撑高） */
  height: number;
  /** 卡片顶部 y（滚动容器坐标） */
  top: number;
  /** 这一页是否因为超高块而被撑高 */
  oversized: boolean;
}

export interface DocxLayout {
  /** 分页模式的页列表；连续模式为空数组 */
  pages: DocxPage[];
  /** 块 → 顶部 y（滚动容器坐标），两种模式都填 */
  blockTops: Float64Array;
  /** 块 → 所属页下标（0 起）：`PAGE` 域要按"run 在哪一页"取实时页码 */
  blockPages: Int32Array;
  /** 页 → 顶部 y（连续模式为空） */
  pageTops: Float64Array;
  /** 内容总高度（滚动条长度） */
  totalHeight: number;
}

/**
 * 连续模式：块首尾相接（不分页），坐标就是前缀和。
 * 保留它是因为它是"看全文"最省事的形态，也是分页模式的兜底对照。
 */
export function layoutContinuous(heights: Float64Array): DocxLayout {
  const total = heights.length;
  const blockTops = new Float64Array(total);
  let y = 0;
  for (let i = 0; i < total; i += 1) {
    blockTops[i] = y;
    y += heights[i];
  }
  return {
    pages: [],
    blockTops,
    blockPages: new Int32Array(total),
    pageTops: new Float64Array(0),
    totalHeight: y,
  };
}

/**
 * 分页模式：按几何把块装页。
 *
 * `startsPage(i)` 由调用方给出（显式分页符 / 段前分页 / 分节符），
 * `heights` 是**已含缩放**的每块高度（虚拟滚动的估算/测量结果）。
 */
export function layoutPages(
  heights: Float64Array,
  geometry: PageGeometry,
  startsPage: (index: number) => boolean,
): DocxLayout {
  const total = heights.length;
  const blockTops = new Float64Array(total);
  const blockPages = new Int32Array(total);
  if (total === 0) {
    return {
      pages: [],
      blockTops,
      blockPages,
      pageTops: new Float64Array(0),
      totalHeight: 0,
    };
  }

  const contentHeight = geometry.contentHeightPx;
  /**
   * 纸张卡片的名义内容高（**不减页脚**）：卡片高度要正好等于页面高度，
   * 页脚画在下边距里。分页判断仍用 `contentHeight`（少 26px 页脚，保证正文不会压到页码）。
   */
  const nominalContentHeight = Math.max(40, round2(geometry.heightPx - geometry.marginTopPx - geometry.marginBottomPx));
  const pages: DocxPage[] = [];
  let start = 0;
  let inside = 0; // 当前页已用内容高

  const pushPage = (endIndex: number, top: number): number => {
    // 超高块：这一页按内容撑高（不裁切、不与下一页重叠）
    const oversized = inside > contentHeight + 0.5;
    const contentBox = Math.max(nominalContentHeight, inside);
    const height = round2(geometry.marginTopPx + contentBox + geometry.marginBottomPx);
    pages.push({
      index: pages.length,
      start,
      end: endIndex,
      contentHeight: round2(inside),
      height,
      top: round2(top),
      oversized,
    });
    return top + height + PAGE_GAP;
  };

  let top = 0;
  for (let i = 0; i < total; i += 1) {
    const height = heights[i];
    if (i > start) {
      const explicit = startsPage(i);
      // 放不下就自然分页；超高块（比整页还高）也必须换页，否则会把这一页撑到离谱
      const overflow = inside + height > contentHeight + 0.5;
      const tooTall = height > contentHeight + 0.5 && inside > 0;
      if (explicit || overflow || tooTall) {
        top = pushPage(i - 1, top);
        start = i;
        inside = 0;
      }
    }
    blockTops[i] = round2(top + geometry.marginTopPx + inside);
    // 块 → 页：`PAGE` 域要按"这个 run 落在第几页"取实时页码
    blockPages[i] = pages.length;
    inside += height;
  }
  top = pushPage(total - 1, top);

  const pageTops = new Float64Array(pages.length);
  pages.forEach((page, index) => {
    pageTops[index] = page.top;
  });
  return { pages, blockTops, blockPages, pageTops, totalHeight: Math.max(0, top - PAGE_GAP) };
}

/** 二分：块坐标数组里找「最后一个 top <= y」的块下标（两种模式通用，blockTops 单调不减） */
export function blockIndexAt(blockTops: Float64Array, y: number): number {
  const total = blockTops.length;
  if (total === 0) return 0;
  if (y <= blockTops[0]) return 0;
  let lo = 0;
  let hi = total - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (blockTops[mid] <= y) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** 二分：页坐标里找第一个与 `top` 之后的内容相交的页（返回不下标越界的下标） */
export function firstPageIntersecting(pages: DocxPage[], top: number): number {
  if (pages.length === 0) return 0;
  let lo = 0;
  let hi = pages.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    const page = pages[mid];
    if (page.top + page.height < top) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}
