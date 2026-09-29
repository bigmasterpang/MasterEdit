/**
 * 块列表 + 虚拟滚动。
 *
 * ── 为什么难 ──────────────────────────────────────────────────────────
 * 表格的行高是固定的，段落的**高度取决于换行**，而换行只有浏览器知道。
 * 所以走「先估算 → 渲染后测量修正 → 滚动锚点补偿」三步（思路与 SheetGrid 的
 * 稀疏行高索引一致，但这里块高差异更大：一个段落 18px，一个表格可能 800px）。
 *
 * ── 高度索引 ──────────────────────────────────────────────────────────
 * `heightOf(i) = 测量值 ?? 估算值 ?? 默认值`，用一个**稠密前缀和**支撑三个操作：
 *   · topOf(i)：块 i 的顶部 y（绝对定位用）；
 *   · indexAt(y)：y 处是第几块（滚动 → 渲染窗口）；
 *   · totalHeight：滚动条总高。
 * 只有「测量值集合 / 缩放 / 内容宽度 / 总块数 / 缓存版本」变化才重建；**滚动本身绝不重建**。
 *
 * ── 测量修正 ──────────────────────────────────────────────────────────
 * 渲染后读每块外层的 offsetHeight，与索引里的值不同就记进 `measured`，并按帧合并
 * （一帧内多次测量只重建一次索引，避免滚动时每块都触发 O(总块数) 的重建）。
 * 高度变化会让下方内容整体平移，所以重建后按「视口顶部那一块」补偿 scrollTop，
 * 用户看到的内容不会跳。测试环境（jsdom）offsetHeight 恒为 0，直接忽略测量值。
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { DocBlock } from "../../types";
import { DocxBlock } from "./DocxBlock";
import {
  DEFAULT_BLOCK_HEIGHT,
  PAGE_PADDING,
  PAGE_WIDTH,
  estimateBlockHeight,
  round2,
} from "./docxStyle";
import type { DocxRenderContext } from "./docxRender";
import type { DocxBlocksApi } from "./useDocxBlocks";

/** 渲染窗口上下各多渲染几块（滚动时不会看到空白） */
const OVERSCAN = 6;
/** 窗口内至少渲染这么多块（估算偏差大时也不至于出现空白屏） */
const MIN_RENDER = 10;
/** 一次最多渲染多少块（防止某些块估算高度极小导致窗口巨大） */
const MAX_RENDER = 160;

/** 按帧调度的通用工具（测试环境可能没有 requestAnimationFrame） */
function scheduleFrame(callback: () => void): number {
  if (typeof requestAnimationFrame === "function") return requestAnimationFrame(callback);
  return window.setTimeout(callback, 16) as unknown as number;
}

/* ============================== 高度索引 ============================== */

interface HeightIndex {
  /** 总内容高度（px） */
  totalHeight: number;
  /** 块 i 的顶部 y */
  topOf(index: number): number;
  /** y 落在第几块（0..total-1） */
  indexAt(y: number): number;
}

/**
 * 稠密前缀和。块数是**顶层块数**（文档级通常几百 ~ 几万），重建一次 O(n) 是微秒级，
 * 不值得像百万行工作表那样上稀疏表。
 */
function buildHeightIndex(total: number, heightOf: (index: number) => number): HeightIndex {
  const prefix = new Float64Array(total + 1);
  for (let i = 0; i < total; i += 1) prefix[i + 1] = prefix[i] + heightOf(i);
  const totalHeight = prefix[total];

  return {
    totalHeight,
    topOf(index: number): number {
      if (index <= 0) return 0;
      if (index >= total) return totalHeight;
      return prefix[index];
    },
    indexAt(y: number): number {
      if (total <= 0) return 0;
      if (y <= 0) return 0;
      if (y >= totalHeight) return total - 1;
      // 二分找最后一个 prefix[i] <= y
      let lo = 0;
      let hi = total;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (prefix[mid] <= y) lo = mid + 1;
        else hi = mid;
      }
      return Math.min(total - 1, Math.max(0, lo - 1));
    },
  };
}

/** 默认视口（还没量到真实尺寸时用；jsdom 下 clientHeight 恒为 0，也靠它兜底） */
const FALLBACK_VIEWPORT = { width: PAGE_WIDTH, height: 600 };

/* ============================== 组件 ============================== */

export interface DocxBlocksProps {
  api: DocxBlocksApi;
  ctx: DocxRenderContext;
  /** 跳转请求：token 变化就滚到 index（查找命中 / 大纲点击用） */
  scrollRequest: { index: number; token: number } | null;
}

export function DocxBlocks({ api, ctx, scrollRequest }: DocxBlocksProps) {
  const scrollerRef = useRef<HTMLDivElement | null>(null);
  const pageRef = useRef<HTMLDivElement | null>(null);
  const [viewport, setViewport] = useState(FALLBACK_VIEWPORT);
  const [scrollTop, setScrollTop] = useState(0);
  const [heightVersion, setHeightVersion] = useState(0);
  const mountedRef = useRef(true);

  /* ---------- 测量值 ---------- */
  const measuredRef = useRef(new Map<number, number>());
  const measurePendingRef = useRef(false);

  /* ---------- 估算缓存（每个组件实例一份，跟随「缩放 + 内容宽 + 缓存版本」失效） ---------- */
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

  /* ---------- 页面宽 & 可用内容宽（估算与渲染必须用同一个值） ---------- */
  const layout = useMemo(() => {
    const pageWidth = Math.max(320, Math.min(viewport.width - 16, PAGE_WIDTH));
    return { pageWidth, contentWidth: Math.max(160, pageWidth - PAGE_PADDING * 2) };
  }, [viewport.width]);

  /* ---------- 高度：测量值优先，其次估算（带缓存） ---------- */
  const blockAt = api.blockAt;
  const heightOf = useCallback(
    (index: number): number => {
      const measured = measuredRef.current.get(index);
      if (measured !== undefined && measured > 0) return measured;
      const block = blockAt(index);
      if (!block) return DEFAULT_BLOCK_HEIGHT;
      const key = `${ctx.scale}|${layout.contentWidth}|${api.version}`;
      if (estimateRef.current.key !== key) estimateRef.current = { key, map: new Map() };
      const cached = estimateRef.current.map.get(index);
      if (cached !== undefined) return cached;
      const height = estimateBlockHeight(block, ctx.scale, layout.contentWidth);
      estimateRef.current.map.set(index, height);
      return height;
    },
    [blockAt, ctx.scale, layout.contentWidth, api.version],
  );

  const total = api.total;
  const index = useMemo(() => buildHeightIndex(total, heightOf), [total, heightOf, heightVersion]);

  /* ---------- 可见窗口 ---------- */
  const range = useMemo(() => {
    if (total <= 0) return { first: 0, last: -1 };
    const firstVisible = index.indexAt(scrollTop);
    const lastVisible = index.indexAt(scrollTop + viewport.height);
    let first = Math.max(0, firstVisible - OVERSCAN);
    let last = Math.min(total - 1, lastVisible + OVERSCAN);
    // 估算普遍偏小时窗口会太窄：补齐到最小渲染块数，避免出现空白屏
    while (last - first + 1 < MIN_RENDER && (first > 0 || last < total - 1)) {
      if (last < total - 1) last += 1;
      if (last - first + 1 >= MIN_RENDER) break;
      if (first > 0) first -= 1;
    }
    if (last - first + 1 > MAX_RENDER) last = first + MAX_RENDER - 1;
    return { first, last };
  }, [index, scrollTop, total, viewport.height]);

  /* ---------- 取数：确保窗口内的块都在缓存里 ---------- */
  useEffect(() => {
    if (total <= 0 || range.last < range.first) return;
    api.ensure(range.first, range.last);
  }, [api, range.first, range.last, total]);

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
    const page = pageRef.current;
    if (!page) return;
    let changed = false;
    for (const child of Array.from(page.children)) {
      const raw = child.getAttribute("data-docx-block-index");
      if (raw === null) continue;
      const blockIndex = Number(raw);
      if (!Number.isFinite(blockIndex)) continue;
      const height = (child as HTMLElement).offsetHeight;
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
    const page = pageRef.current;
    if (!page || typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(() => measurePass());
    observer.observe(page);
    return () => observer.disconnect();
  }, [measurePass]);

  /* ---------- 高度变化后的滚动锚点补偿 ---------- */
  const anchorRef = useRef<{ index: HeightIndex; scrollTop: number } | null>(null);
  useLayoutEffect(() => {
    const el = scrollerRef.current;
    const previous = anchorRef.current;
    if (el && previous && previous.index !== index) {
      // 以「视口顶部那一块」为锚：它的顶部 y 平移了多少，scrollTop 就补多少
      const anchorIndex = previous.index.indexAt(previous.scrollTop);
      const delta = index.topOf(anchorIndex) - previous.index.topOf(anchorIndex);
      if (delta !== 0) {
        el.scrollTop = Math.max(0, el.scrollTop + delta);
        setScrollTop(el.scrollTop);
      }
    }
    if (el) anchorRef.current = { index, scrollTop: el.scrollTop };
  }, [index]);

  /* ---------- 跳转（查找命中 / 大纲点击） ---------- */
  const handledTokenRef = useRef<number | null>(null);
  useLayoutEffect(() => {
    if (!scrollRequest || handledTokenRef.current === scrollRequest.token) return;
    handledTokenRef.current = scrollRequest.token;
    const el = scrollerRef.current;
    if (!el || total <= 0) return;
    const target = Math.max(0, Math.min(total - 1, scrollRequest.index));
    // 留一点上边距，命中的块不要贴着容器顶
    el.scrollTop = Math.max(0, index.topOf(target) - 8);
    setScrollTop(el.scrollTop);
  }, [scrollRequest, index, total]);

  /* ---------- 渲染 ---------- */
  const rendered = useMemo(() => {
    const list: Array<{ index: number; block: DocBlock; top: number }> = [];
    if (total <= 0 || range.last < range.first) return list;
    for (let i = range.first; i <= range.last; i += 1) {
      const block = api.blockAt(i);
      if (block) list.push({ index: i, block, top: index.topOf(i) });
    }
    return list;
    // api.version：缓存里来了新块要重算；index：高度索引换了要重排
  }, [api, range.first, range.last, index, total]);

  const isLoadingFirst = api.loading && total === 0;

  return (
    <div
      ref={scrollerRef}
      data-docx-scroll="true"
      onScroll={onScroll}
      className="relative h-full overflow-auto bg-app"
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
          ref={pageRef}
          data-docx-page="true"
          style={{
            position: "relative",
            width: `${layout.pageWidth}px`,
            height: `${round2(index.totalHeight)}px`,
            margin: "0 auto",
            paddingLeft: `${PAGE_PADDING}px`,
            paddingRight: `${PAGE_PADDING}px`,
            boxSizing: "border-box",
          }}
        >
          {rendered.map((entry) => (
            <div
              key={entry.index}
              data-docx-block="true"
              data-docx-block-index={entry.index}
              data-docx-block-kind={entry.block.kind}
              data-docx-hit={ctx.highlightBlock === entry.index ? "true" : undefined}
              style={{
                position: "absolute",
                top: `${round2(entry.top)}px`,
                // 绝对定位把块限制在页面留白之间，与估算用的 contentWidth 完全一致
                left: `${PAGE_PADDING}px`,
                right: `${PAGE_PADDING}px`,
                backgroundColor:
                  ctx.highlightBlock === entry.index ? "rgba(255, 214, 102, 0.28)" : undefined,
                borderRadius: ctx.highlightBlock === entry.index ? "3px" : undefined,
              }}
            >
              <DocxBlock block={entry.block} ctx={ctx} depth={0} />
            </div>
          ))}
          {rendered.length === 0 ? (
            <div
              data-docx-block-placeholder="true"
              style={{ position: "absolute", top: 0, left: PAGE_PADDING, right: PAGE_PADDING }}
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
