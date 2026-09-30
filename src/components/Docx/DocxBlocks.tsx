/**
 * 块列表 + 虚拟滚动（**分页视图 / 连续视图**两种模式）。
 *
 * ── 为什么难 ──────────────────────────────────────────────────────────
 * 表格的行高是固定的，段落的**高度取决于换行**，而换行只有浏览器知道。
 * 所以走「先估算 → 渲染后测量修正 → 滚动锚点补偿」三步：
 * 估算见 `docxStyle.ts`（逐字符扫行数 + 行高 + 段前段后），测量见下面的 `measurePass`。
 *
 * ── 两种模式共用同一套坐标 ─────────────────────────────────────────────
 * `docxPages.ts` 把「每块高度」换算成 **块 → 滚动容器内 y**：
 *   · `paged`：先按页几何把块装进一页一页（显式分页符强制换页、放不下自然换页、
 *     超高块独占一页并把卡片撑高），块坐标 = 页卡片 top + 页上边距 + 页内 y；
 *   · `continuous`：块首尾相接（老行为，一行到底）。
 * 两种模式都产出 `blockTops`（单调不减），于是「可见窗口 / 跳转 / 锚点补偿 / 测量」
 * 全都是同一套代码 —— 这也保证了分页后**不会**退化成全量渲染。
 *
 * ── 虚拟滚动 ──────────────────────────────────────────────────────────
 * 分页模式按**页**虚拟化（页是天然的大粒度单元：A4 一页约 1100px，一屏放不下两页），
 * 只渲染与视口相交的页 + 上下各 1 页缓冲；连续模式仍按块虚拟化。
 * 两种模式的「合并后的同类占位块」都只渲染一张卡片（见 `planBlocks`）。
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from "react";
import type { DocBlock } from "../../types";
import { DocxBlock, UnsupportedCard } from "./DocxBlock";
import {
  DEFAULT_BLOCK_HEIGHT,
  PAGE_PADDING,
  PAGE_WIDTH,
  estimateBlockHeight,
  estimateUnsupportedHeight,
  round2,
} from "./docxStyle";
import { planBlocks, type BlockPlan } from "./docxGrid";
import {
  PAGE_FOOTER_HEIGHT,
  blockIndexAt,
  firstPageIntersecting,
  layoutContinuous,
  layoutPages,
  resolvePageGeometry,
  type DocxLayout,
  type DocxPage,
  type PageGeometry,
} from "./docxPages";
import type { DocxRenderContext } from "./docxRender";
import type { DocxBlocksApi } from "./useDocxBlocks";

/** 渲染窗口上下各多渲染几块 / 几页（滚动时不会看到空白） */
const OVERSCAN = 6;
/** 分页模式：可见页上下各多渲染几页 */
const PAGE_OVERSCAN = 1;
/** 窗口内至少渲染这么多块（估算偏差大时也不至于出现空白屏） */
const MIN_RENDER = 10;
/** 一次最多渲染多少块（防止某些块估算高度极小导致窗口巨大） */
const MAX_RENDER = 160;
/** 背景画布上下留白（px）：纸张不要贴着滚动条 */
const BOOK_PADDING_Y = 16;
/**
 * 画布留白以**坐标偏移**的方式加入（而不是 canvas 的 padding）：
 * 绝对定位子元素的 top 是相对父元素 padding box 的，用 padding 会让"布局坐标"与
 * 实际位置差一个 padding —— 跳转、锚点补偿、页内块定位三处都会跟着歪。
 */
const BOOK_OFFSET = BOOK_PADDING_Y;

export type DocxViewMode = "paged" | "continuous";

/** 按帧调度的通用工具（测试环境可能没有 requestAnimationFrame） */
function scheduleFrame(callback: () => void): number {
  if (typeof requestAnimationFrame === "function") return requestAnimationFrame(callback);
  return window.setTimeout(callback, 16) as unknown as number;
}

/** 默认视口（还没量到真实尺寸时用；jsdom 下 clientHeight 恒为 0，也靠它兜底） */
const FALLBACK_VIEWPORT = { width: PAGE_WIDTH + 32, height: 600 };

/**
 * 窗口比纸张窄时按比例整体缩小（等比缩放页面与页边距，等同 Word 的"适应页宽"）。
 * 缩放后版心宽/高一起变，估算与渲染仍用同一组数字。
 */
export function fitPageGeometry(base: PageGeometry, maxWidth: number): PageGeometry {
  if (maxWidth <= 0 || base.widthPx <= maxWidth) return base;
  const factor = maxWidth / base.widthPx;
  const widthPx = round2(base.widthPx * factor);
  const heightPx = round2(base.heightPx * factor);
  const marginTopPx = round2(base.marginTopPx * factor);
  const marginRightPx = round2(base.marginRightPx * factor);
  const marginBottomPx = round2(base.marginBottomPx * factor);
  const marginLeftPx = round2(base.marginLeftPx * factor);
  return {
    widthPx,
    heightPx,
    marginTopPx,
    marginRightPx,
    marginBottomPx,
    marginLeftPx,
    contentWidthPx: Math.max(120, round2(widthPx - marginLeftPx - marginRightPx)),
    contentHeightPx: Math.max(
      80,
      round2(heightPx - marginTopPx - marginBottomPx - PAGE_FOOTER_HEIGHT),
    ),
  };
}

export interface DocxBlocksProps {
  api: DocxBlocksApi;
  ctx: DocxRenderContext;
  /** 分页（一页一页的纸）还是连续（一条流） */
  mode: DocxViewMode;
  /** 跳转请求：token 变化就滚到 index（查找命中 / 大纲点击用） */
  scrollRequest: { index: number; token: number } | null;
}

export function DocxBlocks({ api, ctx, mode, scrollRequest }: DocxBlocksProps) {
  const scrollerRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLDivElement | null>(null);
  const [viewport, setViewport] = useState(FALLBACK_VIEWPORT);
  const [scrollTop, setScrollTop] = useState(0);
  const [heightVersion, setHeightVersion] = useState(0);
  const mountedRef = useRef(true);

  /* ---------- 测量值 ---------- */
  const measuredRef = useRef(new Map<number, number>());
  const measurePendingRef = useRef(false);

  /* ---------- 估算缓存（跟随「缩放 + 内容宽 + 缓存版本」失效） ---------- */
  const estimateRef = useRef<{ key: string; map: Map<number, number> }>({ key: "", map: new Map() });

  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
    };
  }, []);

  /* ---------- 视口尺寸：ResizeObserver + 首次布局读取 ---------- */
  const readViewport = useCallback(() => {
    const el = scrollerRef.current;
    if (!el) return;
    const width = el.clientWidth || FALLBACK_VIEWPORT.width;
    const height = el.clientHeight || FALLBACK_VIEWPORT.height;
    setViewport((prev) => (prev.width === width && prev.height === height ? prev : { width, height }));
    setScrollTop(el.scrollTop);
  }, []);

  useLayoutEffect(() => {
    readViewport();
  }, [readViewport]);

  /**
   * 把滚动容器交给渲染上下文：图片懒加载的 IntersectionObserver 要拿它当 root
   * （用 viewport 当 root 时容器 overflow 会裁掉目标，200px 预取范围失效）。
   * 放在 layout 阶段同步：子组件的 `useEffect` 一定在父组件 layout 效果之后才跑。
   */
  useLayoutEffect(() => {
    ctx.scrollRootRef.current = scrollerRef.current;
    return () => {
      ctx.scrollRootRef.current = null;
    };
  });

  useEffect(() => {
    const el = scrollerRef.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(() => readViewport());
    observer.observe(el);
    return () => observer.disconnect();
  }, [readViewport]);

  /* ---------- 滚动：按帧节流，避免每一像素都重算窗口 ---------- */
  const scrollRafRef = useRef<number | null>(null);
  const onScroll = useCallback(() => {
    if (scrollRafRef.current !== null) return;
    scrollRafRef.current = scheduleFrame(() => {
      scrollRafRef.current = null;
      const el = scrollerRef.current;
      if (el) setScrollTop(el.scrollTop);
    });
  }, []);

  useEffect(
    () => () => {
      if (scrollRafRef.current !== null && typeof cancelAnimationFrame === "function") {
        cancelAnimationFrame(scrollRafRef.current);
      }
    },
    [],
  );

  /* ---------- 页面几何：后端 sectPr → 兜底 A4 → 窗口太窄时等比缩小 ---------- */
  const geometry = useMemo(
    () => fitPageGeometry(resolvePageGeometry(api.page), Math.max(240, viewport.width - 24)),
    [api.page, viewport.width],
  );

  /* ---------- 可用内容宽：两种模式各自的版心（估算与渲染必须用同一个值） ---------- */
  const contentWidth = useMemo(() => {
    if (mode === "paged") return geometry.contentWidthPx;
    const paper = Math.max(320, Math.min(viewport.width - 16, PAGE_WIDTH));
    return Math.max(160, paper - PAGE_PADDING * 2);
  }, [mode, geometry.contentWidthPx, viewport.width]);

  /* ---------- 每块高度：测量值优先，其次估算（合并组按整张卡片算一次） ---------- */
  const blockAt = api.blockAt;
  const heightOf = useCallback(
    (index: number, plan: BlockPlan): number => {
      if (plan.count[index] === 0) return 0; // 合并组内的非首块：并进首块那张卡片
      const measured = measuredRef.current.get(index);
      if (measured !== undefined && measured > 0) return measured;
      const key = `${ctx.scale}|${contentWidth}|${api.version}`;
      if (estimateRef.current.key !== key) estimateRef.current = { key, map: new Map() };
      const cached = estimateRef.current.map.get(index);
      if (cached !== undefined) return cached;
      let height: number;
      const count = plan.count[index];
      if (count > 1) {
        // 合并卡片：标题行 + 组内每条说明各按 12px 字号估行
        const details: string[] = [];
        for (let member = index; member < index + count; member += 1) {
          const block = blockAt(member);
          if (block && block.kind === "unsupported") details.push(block.detail);
        }
        height = estimateUnsupportedHeight(details, contentWidth);
      } else {
        const block = blockAt(index);
        height = block ? estimateBlockHeight(block, ctx.scale, contentWidth) : DEFAULT_BLOCK_HEIGHT;
      }
      estimateRef.current.map.set(index, height);
      return height;
    },
    [blockAt, ctx.scale, contentWidth, api.version],
  );

  const total = api.total;
  /** 每块高度（合并组成员记 0）+ 分页 / 合并计划：任一变化都要重排 */
  const plan = useMemo(() => planBlocks(total, blockAt), [total, blockAt, api.version]);
  const layout: DocxLayout = useMemo(() => {
    const heights = new Float64Array(total);
    for (let index = 0; index < total; index += 1) heights[index] = heightOf(index, plan);
    // heightVersion 必须在依赖里：测量值存在 ref 里，只有它变化才知道要重排
    return mode === "paged"
      ? layoutPages(heights, geometry, (index) => plan.startsPage[index] === 1)
      : layoutContinuous(heights);
  }, [total, plan, heightOf, heightVersion, mode, geometry]);

  /* ---------- 可见窗口 ---------- */
  const visibleWindow = useMemo(() => {
    if (total <= 0) return { first: 0, last: -1, pages: [] as DocxPage[] };
    // 滚动坐标 → 布局坐标（差一个画布留白）
    const contentTop = Math.max(0, scrollTop - BOOK_OFFSET);
    if (mode === "paged") {
      const pages = layout.pages;
      if (pages.length === 0) return { first: 0, last: -1, pages: [] as DocxPage[] };
      const startPage = Math.max(0, firstPageIntersecting(pages, contentTop) - PAGE_OVERSCAN);
      const visible: DocxPage[] = [];
      const limit = contentTop + viewport.height;
      for (let index = startPage; index < pages.length; index += 1) {
        const page = pages[index];
        visible.push(page);
        // 多渲染一页缓冲就够（一页 ≈ 1100px，缓冲再多就是白渲染）
        if (page.top > limit && visible.length > PAGE_OVERSCAN + 1) break;
        if (visible.length > 8) break;
      }
      const first = visible.length > 0 ? visible[0].start : 0;
      const last = visible.length > 0 ? visible[visible.length - 1].end : -1;
      return { first, last, pages: visible };
    }
    const firstVisible = blockIndexAt(layout.blockTops, contentTop);
    const lastVisible = blockIndexAt(layout.blockTops, contentTop + viewport.height);
    let first = Math.max(0, firstVisible - OVERSCAN);
    let last = Math.min(total - 1, lastVisible + OVERSCAN);
    // 估算普遍偏小时窗口会太窄：补齐到最小渲染块数，避免出现空白屏
    while (last - first + 1 < MIN_RENDER && (first > 0 || last < total - 1)) {
      if (last < total - 1) last += 1;
      if (last - first + 1 >= MIN_RENDER) break;
      if (first > 0) first -= 1;
    }
    if (last - first + 1 > MAX_RENDER) last = first + MAX_RENDER - 1;
    return { first, last, pages: [] as DocxPage[] };
  }, [layout, mode, scrollTop, total, viewport.height]);

  /* ---------- 取数：确保窗口内的块都在缓存里 ---------- */
  useEffect(() => {
    if (total <= 0 || visibleWindow.last < visibleWindow.first) return;
    api.ensure(visibleWindow.first, visibleWindow.last);
  }, [api, visibleWindow.first, visibleWindow.last, total]);

  /* ---------- 测量：渲染后读实际高度，按帧合并回写 ---------- */
  const flushMeasurements = useCallback(() => {
    if (measurePendingRef.current) return;
    measurePendingRef.current = true;
    scheduleFrame(() => {
      measurePendingRef.current = false;
      if (!mountedRef.current) return;
      setHeightVersion((value) => value + 1);
    });
  }, []);

  const measurePass = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    let changed = false;
    for (const element of Array.from(canvas.querySelectorAll<HTMLElement>("[data-docx-block-index]"))) {
      const raw = element.getAttribute("data-docx-block-index");
      if (raw === null) continue;
      const blockIndex = Number(raw);
      if (!Number.isFinite(blockIndex)) continue;
      const height = element.offsetHeight;
      // jsdom 里 offsetHeight 恒为 0；真实浏览器里 0 也表示"还没布局"，都要忽略
      if (!(height > 0)) continue;
      const previous = measuredRef.current.get(blockIndex);
      if (previous !== undefined && Math.abs(previous - height) < 0.5) continue;
      measuredRef.current.set(blockIndex, height);
      changed = true;
    }
    if (changed) flushMeasurements();
  }, [flushMeasurements]);

  // 每次渲染后测一遍：窗口只有几十块，读 offsetHeight 的代价可忽略
  useLayoutEffect(() => {
    measurePass();
  });

  // 图片异步加载、字体替换会改变块高：用 ResizeObserver 兜住这两类「渲染后还会变」的情况
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(() => measurePass());
    observer.observe(canvas);
    return () => observer.disconnect();
  }, [measurePass]);

  /* ---------- 高度变化后的滚动锚点补偿 ---------- */
  const anchorRef = useRef<{ tops: Float64Array; totalHeight: number; scrollTop: number } | null>(null);
  useLayoutEffect(() => {
    const el = scrollerRef.current;
    const previous = anchorRef.current;
    if (el && previous && previous.tops !== layout.blockTops) {
      // 以「视口顶部那一块」为锚：它的顶部 y 平移了多少，scrollTop 就补多少
      const anchorIndex = blockIndexAt(previous.tops, previous.scrollTop);
      const nextTops = layout.blockTops;
      const shared = Math.min(anchorIndex, nextTops.length - 1, previous.tops.length - 1);
      if (shared >= 0) {
        const delta = nextTops[shared] - previous.tops[shared];
        if (delta !== 0) {
          el.scrollTop = Math.max(0, el.scrollTop + delta);
          setScrollTop(el.scrollTop);
        }
      }
    }
    if (el) anchorRef.current = { tops: layout.blockTops, totalHeight: layout.totalHeight, scrollTop: el.scrollTop };
  }, [layout]);

  /* ---------- 跳转（查找命中 / 大纲点击） ---------- */
  const handledTokenRef = useRef<number | null>(null);
  useLayoutEffect(() => {
    if (!scrollRequest || handledTokenRef.current === scrollRequest.token) return;
    handledTokenRef.current = scrollRequest.token;
    const el = scrollerRef.current;
    if (!el || total <= 0) return;
    const target = Math.max(0, Math.min(total - 1, scrollRequest.index));
    const y = target < layout.blockTops.length ? layout.blockTops[target] + BOOK_OFFSET : BOOK_OFFSET;
    // 留一点上边距，命中的块不要贴着容器顶
    el.scrollTop = Math.max(0, y - 8);
    setScrollTop(el.scrollTop);
  }, [scrollRequest, layout, total]);

  /* ---------- 渲染 ---------- */
  const renderedBlocks = useMemo(() => {
    const list: Array<{ index: number; block: DocBlock; top: number; page: number }> = [];
    if (total <= 0 || visibleWindow.last < visibleWindow.first) return list;
    if (mode === "paged") {
      for (const page of visibleWindow.pages) {
        for (let index = page.start; index <= page.end; index += 1) {
          const block = api.blockAt(index);
          if (block) list.push({ index, block, top: layout.blockTops[index] - page.top, page: page.index });
        }
      }
      return list;
    }
    for (let index = visibleWindow.first; index <= visibleWindow.last; index += 1) {
      const block = api.blockAt(index);
      if (block) list.push({ index, block, top: layout.blockTops[index], page: -1 });
    }
    return list;
  }, [api, visibleWindow, layout, mode, total]);

  /** 一个块的外壳：绝对定位 + 查找高亮；合并组的非首块不渲染 */
  const renderShell = (entry: { index: number; block: DocBlock; top: number }): ReactNode => {
    const count = plan.count[entry.index];
    if (count === 0) return null; // 组内非首块：并进上一张卡片
    const highlighted = ctx.highlightBlock === entry.index;
    const grouped = count > 1;
    let content: ReactNode;
    if (grouped) {
      const details: string[] = [];
      for (let member = entry.index; member < entry.index + count; member += 1) {
        const block = api.blockAt(member);
        if (block && block.kind === "unsupported") details.push(block.detail);
      }
      content = (
        <UnsupportedCard
          label={entry.block.kind === "unsupported" ? entry.block.label : ""}
          details={details}
        />
      );
    } else {
      content = <DocxBlock block={entry.block} ctx={ctx} depth={0} />;
    }
    return (
      <div
        key={entry.index}
        data-docx-block="true"
        data-docx-block-index={entry.index}
        data-docx-block-kind={entry.block.kind}
        data-docx-block-count={grouped ? count : undefined}
        data-docx-hit={highlighted ? "true" : undefined}
        style={{
          position: "absolute",
          top: `${round2(entry.top + (mode === "continuous" ? BOOK_OFFSET : 0))}px`,
          // 绝对定位把块限制在版心内，与估算用的 contentWidth 完全一致
          left: mode === "paged" ? `${geometry.marginLeftPx}px` : `${PAGE_PADDING}px`,
          right: mode === "paged" ? `${geometry.marginRightPx}px` : `${PAGE_PADDING}px`,
          backgroundColor: highlighted ? "rgba(255, 214, 102, 0.28)" : undefined,
          borderRadius: highlighted ? "3px" : undefined,
        }}
      >
        {content}
      </div>
    );
  };

  const isLoadingFirst = api.loading && total === 0;
  const pageCount = layout.pages.length;

  return (
    <div
      ref={scrollerRef}
      data-docx-scroll="true"
      data-docx-mode={mode}
      onScroll={onScroll}
      className="relative h-full select-text overflow-auto bg-app"
      style={{ contain: "content" }}
    >
      {isLoadingFirst ? (
        <div className="flex h-full items-center justify-center text-[12px] text-muted">
          正在读取文档内容…
        </div>
      ) : total === 0 ? (
        <div className="flex h-full items-center justify-center text-[12px] text-muted">
          这个文档没有可显示的内容。
        </div>
      ) : (
        <div
          ref={canvasRef}
          data-docx-page={mode === "continuous" ? "true" : undefined}
          data-docx-book={mode === "paged" ? "true" : undefined}
          data-docx-content-width={round2(contentWidth)}
          data-docx-page-content-height={round2(geometry.contentHeightPx)}
          data-docx-page-margin-top={round2(geometry.marginTopPx)}
          data-docx-page-margin-bottom={round2(geometry.marginBottomPx)}
          data-docx-page-width={geometry.widthPx}
          data-docx-page-height={geometry.heightPx}
          style={{
            position: "relative",
            width:
              mode === "paged"
                ? "100%"
                : `${Math.max(320, Math.min(viewport.width - 16, PAGE_WIDTH))}px`,
            // 画布高度 = 内容高 + 上下留白；留白用坐标偏移实现，不用 padding（见 BOOK_OFFSET 注释）
            height: `${round2(layout.totalHeight + BOOK_OFFSET * 2)}px`,
            margin: "0 auto",
            paddingLeft: mode === "paged" ? undefined : `${PAGE_PADDING}px`,
            paddingRight: mode === "paged" ? undefined : `${PAGE_PADDING}px`,
            boxSizing: "border-box",
          }}
        >
          {mode === "paged"
            ? visibleWindow.pages.map((page) => (
                <div
                  key={page.index}
                  data-docx-page="true"
                  data-docx-page-number={page.index + 1}
                  data-docx-page-top={round2(page.top + BOOK_OFFSET)}
                  data-docx-page-height={page.height}
                  data-docx-oversized={page.oversized ? "true" : undefined}
                  className="border border-line bg-panel shadow-sm"
                  style={{
                    position: "absolute",
                    top: `${round2(page.top + BOOK_OFFSET)}px`,
                    left: "50%",
                    transform: "translateX(-50%)",
                    width: `${geometry.widthPx}px`,
                    height: `${page.height}px`,
                  }}
                >
                  {renderedBlocks
                    .filter((entry) => entry.page === page.index)
                    .map((entry) => renderShell(entry))}
                  <div
                    data-docx-page-footer={page.index + 1}
                    className="text-[11px] text-faint"
                    style={{
                      position: "absolute",
                      left: 0,
                      right: 0,
                      bottom: `${round2(Math.max(2, geometry.marginBottomPx / 3))}px`,
                      height: `${PAGE_FOOTER_HEIGHT}px`,
                      lineHeight: `${PAGE_FOOTER_HEIGHT}px`,
                      textAlign: "center",
                    }}
                  >
                    第 {page.index + 1} 页 · 共 {pageCount} 页
                    {page.oversized ? "（本页内容超出一页，未裁切）" : ""}
                  </div>
                </div>
              ))
            : renderedBlocks.map((entry) => renderShell(entry))}
          {renderedBlocks.length === 0 ? (
            <div
              data-docx-block-placeholder="true"
              style={{
                position: "absolute",
                top: BOOK_OFFSET,
                left: mode === "paged" ? geometry.marginLeftPx : PAGE_PADDING,
                right: mode === "paged" ? geometry.marginRightPx : PAGE_PADDING,
              }}
              className="pt-2 text-center text-[11px] text-faint"
            >
              正在取这一段的块…
            </div>
          ) : null}
        </div>
      )}
    </div>
  );
}
