/**
 * 单个块的渲染：段落（含 run 级格式、列表前缀、四边边框）、图片（懒加载）、
 * 形状（线框）、分页提示线、占位卡片。
 *
 * 只读原则（见 `docs/plan-docx.md`）：**不做 contentEditable、不做输入框**，
 * 文字保持可选中复制（不加 `user-select: none`）。
 */
import { memo, useEffect, useRef, useState, type CSSProperties, type ReactNode } from "react";
import type { DocBlock, DocImage, DocParagraph, DocShape, DocTextBox } from "../../types";
import {
  BLOCK_MARGIN_Y,
  DEFAULT_LINE_WIDTH_PT,
  PAGE_BREAK_HEIGHT,
  TEXT_BOX_DEFAULT_BORDER_PT,
  TEXT_BOX_PADDING,
  borderStyleToCss,
  breakHintText,
  estimateShapeHeight,
  fieldRunMinWidthPx,
  hexColor,
  listGapPx,
  MEDIA_RETRY_MS,
  MEDIA_TIMEOUT_MS,
  paragraphBoxStyle,
  ptToPx,
  resolveImageBox,
  resolveRunText,
  round2,
  runInlineStyle,
  spacerFontFamilies,
  splitRunSegments,
} from "./docxStyle";
import { DocxTable } from "./DocxTable";
import { mediaExtension, mediaKind, type DocxRenderContext } from "./docxRender";

/** 淡色的分页/分节提示（不强行分页，只提示原来的分页位置） */
const HINT_LINE = "1px dashed rgba(140, 148, 164, 0.75)";
const HINT_COLOR = "rgb(140, 148, 164)";
const PLACEHOLDER_BORDER = "1px dashed rgba(140, 148, 164, 0.7)";
const PLACEHOLDER_BG = "rgba(127, 140, 160, 0.08)";
/** 图片占位/骨架框的最大显示宽度（px） */
const MEDIA_MAX_WIDTH = 520;

export interface DocxBlockProps {
  block: DocBlock;
  ctx: DocxRenderContext;
  /** 嵌套深度（表格单元格里的块会 +1） */
  depth?: number;
  /**
   * 当前块的**可用内容宽**（px）。图片按它决定显示盒子（`resolveImageBox`），
   * 高度估算用的是同一个值 —— 表格单元格里传的是单元格内宽，不是整页宽。
   */
  contentWidth: number;
}

/* ------------------------------ 分页提示线 ------------------------------ */

function PageBreakLine({ label }: { label: string }) {
  return (
    <div
      data-docx-pagebreak="true"
      style={{
        height: `${PAGE_BREAK_HEIGHT}px`,
        lineHeight: `${PAGE_BREAK_HEIGHT - 4}px`,
        display: "flex",
        alignItems: "center",
        gap: "6px",
        fontSize: "11px",
        color: HINT_COLOR,
        // 提示线自身不参与行距继承，避免把段高算歪
        whiteSpace: "nowrap",
      }}
    >
      <span style={{ flex: "1 1 auto", borderTop: HINT_LINE }} />
      <span>{label}</span>
      <span style={{ flex: "1 1 auto", borderTop: HINT_LINE }} />
    </div>
  );
}

/* -------------------------------- 段落 -------------------------------- */

const DocxParagraph = memo(function DocxParagraph({
  paragraph,
  ctx,
}: {
  paragraph: DocParagraph;
  ctx: DocxRenderContext;
}) {
  const hint = breakHintText(paragraph);
  const boxStyle: CSSProperties = { ...paragraphBoxStyle(paragraph, ctx.scale), position: "relative" };
  /**
   * 子节点用数组拼装，**不能写成 JSX 里的多行元素** —— 段落是 `white-space: pre-wrap`，
   * JSX 元素之间的缩进/换行会变成真实的空白文本节点，直接改变段落内容与换行。
   */
  const children: ReactNode[] = [];
  if (hint) children.push(<PageBreakLine key="break" label={hint} />);
  if (paragraph.list) {
    const baseRun = paragraph.runs[0];
    const baseStyle = baseRun ? runInlineStyle(baseRun, ctx.scale) : {};
    children.push(
      <span
        key="prefix"
        data-docx-list-prefix="true"
        style={{
          display: "inline-block",
          fontFamily: baseStyle.fontFamily,
          fontSize: baseStyle.fontSize,
          paddingRight: `${round2(listGapPx(paragraph, ctx.scale))}px`,
          whiteSpace: "pre",
        }}
      >
        {paragraph.list.prefix}
      </span>,
    );
  }
  paragraph.runs.forEach((run, index) => {
    /**
     * `PAGE` / `NUMPAGES` 域：用**实时值**替换缓存文本（`resolveRunText`），
     * 并按缓存文本宽度给一个 `min-width` + 居中 —— 数字宽度变化不会改变行宽，
     * 也就不会"改分页 → 重新测量 → 再改分页"地来回抖。
     *
     * **判定必须写成"等于 PAGE 或 NUMPAGES"**：老模型里根本没有 `field` 字段
     * （`undefined`），用 `field !== null` 判会把**所有 run** 都当成域 run，
     * 加上 `display: inline-block` 后相邻 run 变成不可断行的原子块 ——
     * 实测后果是整行被 justify 拉成「高速    采样    +    嵌入式开发板」并多占一行。
     */
    const isField = run.field === "PAGE" || run.field === "NUMPAGES";
    const text = resolveRunText(run, ctx.pageNumber, ctx.totalPages);
    const style: CSSProperties = runInlineStyle(run, ctx.scale);
    if (isField) {
      style.display = "inline-block";
      style.minWidth = `${fieldRunMinWidthPx(run, ctx.scale)}px`;
      style.textAlign = "center";
      children.push(
        <span
          key={`run-${index}`}
          data-docx-run="true"
          data-docx-field={run.field ?? undefined}
          style={style}
        >
          {text}
        </span>,
      );
      return;
    }
    /**
     * 普通 run：**把"连续 ≥3 个空白"单独拆成一段**（`splitRunSegments`），
     * 空白段用西文字体渲染 —— Word 逐字符取字体（空格走 `w:ascii`），
     * 宋体的空格是 0.5em、Times New Roman 只有 0.25em，不拆会把整行挤断。
     */
    const segments = splitRunSegments(text);
    if (segments.length <= 1) {
      children.push(
        <span key={`run-${index}`} data-docx-run="true" style={style}>
          {text}
        </span>,
      );
      return;
    }
    const spacerStyle: CSSProperties = { fontFamily: spacerFontFamilies(run).join(", ") };
    segments.forEach((segment, segmentIndex) => {
      children.push(
        segment.spacer ? (
          <span key={`run-${index}-s${segmentIndex}`} data-docx-spacer="true" style={spacerStyle}>
            {segment.text}
          </span>
        ) : (
          <span
            key={`run-${index}-${segmentIndex}`}
            data-docx-run="true"
            data-docx-run-split="true"
            style={style}
          >
            {segment.text}
          </span>
        ),
      );
    });
  });
  // 空段落也要占一行：用零宽空格撑开，避免整段塌掉（`min-height` 只管盒，行盒仍需要内容）
  if (paragraph.runs.length === 0 && !paragraph.list) {
    children.push(
      <span key="empty" data-docx-empty="true">
        {"\u200b"}
      </span>,
    );
  }

  return (
    <div
      data-docx-paragraph="true"
      data-docx-align={paragraph.align ?? undefined}
      data-docx-outline-level={paragraph.outlineLevel !== null ? String(paragraph.outlineLevel) : undefined}
      data-docx-style-id={paragraph.styleId ?? undefined}
      data-docx-list-level={paragraph.list ? String(paragraph.list.level) : undefined}
      style={boxStyle}
    >
      {children}
    </div>
  );
});

/* -------------------------------- 图片 -------------------------------- */

type ImageState =
  | { status: "idle" }
  | { status: "loading" }
  | { status: "ready"; src: string }
  | { status: "error"; message: string };

/** 占位框：显示尺寸与说明文字，失败时给中文原因而不是破图 */
function ImagePlaceholder({
  block,
  message,
  onRetry,
}: {
  block: DocImage;
  message: string;
  /** 失败态给一个「重试」入口：比"等下次进入视口"更直接（超时/后端报错都能自救） */
  onRetry?: () => void;
}) {
  const width = block.widthPx > 0 ? block.widthPx : 220;
  return (
    <div
      data-docx-image-fallback="true"
      style={{
        width: `${round2(Math.min(width, MEDIA_MAX_WIDTH))}px`,
        maxWidth: "100%",
        minHeight: "56px",
        border: PLACEHOLDER_BORDER,
        borderRadius: "6px",
        background: PLACEHOLDER_BG,
        padding: "8px 10px",
        fontSize: "12px",
        color: HINT_COLOR,
        lineHeight: "1.5",
        whiteSpace: "pre-wrap",
      }}
    >
      {message}
      {onRetry ? (
        <div style={{ marginTop: "6px" }}>
          <button
            type="button"
            data-docx-image-retry="true"
            onClick={onRetry}
            style={{
              border: PLACEHOLDER_BORDER,
              borderRadius: "4px",
              background: "transparent",
              padding: "2px 8px",
              fontSize: "12px",
              color: HINT_COLOR,
              cursor: "pointer",
            }}
          >
            重试
          </button>
        </div>
      ) : null}
    </div>
  );
}

const DocxImageBlock = memo(function DocxImageBlock({
  block,
  ctx,
  contentWidth,
}: {
  block: DocImage;
  ctx: DocxRenderContext;
  contentWidth: number;
}) {
  const kind = mediaKind(block.media);
  const cacheKey = ctx.mediaCacheKey(block.media);
  const [state, setState] = useState<ImageState>(() => {
    if (kind === "raster") {
      const cached = ctx.mediaCache.get(cacheKey);
      if (cached) return { status: "ready", src: cached };
    }
    return { status: "idle" };
  });
  const hostRef = useRef<HTMLDivElement | null>(null);
  /** 同一个图片块只请求一次（重复渲染 / 视口反复进出都不会重发）；失败/超时会清掉以便重试 */
  const requestedRef = useRef(false);
  /** 上次发起请求的时间（失败后重试的节流，避免"每次重渲染都重试"打爆后端） */
  const lastAttemptRef = useRef(0);
  /** 当前这次加载效果的"发起请求"函数：失败卡片上的「重试」直接调它，不依赖观测器回调 */
  const requestRef = useRef<(() => void) | null>(null);
  /**
   * ctx 每次渲染都是新对象（逐页带页码），**不能进依赖数组**——
   * 否则每渲染一次就重建一个 IntersectionObserver。用 ref 取最新值即可。
   */
  const ctxRef = useRef(ctx);
  ctxRef.current = ctx;

  useEffect(() => {
    const renderCtx = ctxRef.current;
    if (kind !== "raster") return;
    // 命中缓存（同一张图在前面已经取过）就直接用，不再走 invoke
    const cached = renderCtx.mediaCache.get(cacheKey);
    if (cached) {
      setState({ status: "ready", src: cached });
      return;
    }
    let cancelled = false;
    const request = () => {
      if (requestedRef.current) return;
      // 已失败过一次的，重试要间隔至少 2 秒（避免观测器重建导致的请求风暴）
      const now = Date.now();
      if (lastAttemptRef.current > 0 && now - lastAttemptRef.current < MEDIA_RETRY_MS) return;
      requestedRef.current = true;
      lastAttemptRef.current = now;
      setState({ status: "loading" });
      /**
       * **超时兜底（8 秒）**：后端 `document_media` 要内存解密 + 解包 + base64，
       * 超大图或异常文档上可能长时间不返回。没有超时就会**永远停在「图片加载中…」**
       * （用户实测 `王诚0918.docx`）。超时后走失败态（带中文原因）+ 允许再次进入视口重试。
       */
      let settled = false;
      const timer = window.setTimeout(() => {
        if (settled || cancelled) return;
        settled = true;
        requestedRef.current = false; // 允许重新进入视口后重试
        setState({ status: "error", message: `图片读取超时（超过 ${MEDIA_TIMEOUT_MS / 1000} 秒）` });
      }, MEDIA_TIMEOUT_MS);
      renderCtx
        .loadMedia(block.media)
        .then((src) => {
          if (settled) return;
          settled = true;
          window.clearTimeout(timer);
          if (cancelled) return;
          setState({ status: "ready", src });
        })
        .catch((reason: unknown) => {
          if (settled) return;
          settled = true;
          window.clearTimeout(timer);
          if (cancelled) return;
          // 失败后**清掉"已请求"标记**：再次进入视口要能重试，不能永久卡死
          requestedRef.current = false;
          /**
           * 后端给的是中文原因（字符串）就直接显示；否则把原始错误也带上 ——
           * 只写"图片读取失败"会让用户和排查都没线索。
           */
          const detail =
            typeof reason === "string" && reason.trim()
              ? reason.trim()
              : reason instanceof Error && reason.message
                ? reason.message
                : String(reason);
          setState({ status: "error", message: `图片读取失败：${detail}` });
        });
    };

    const host = hostRef.current;
    requestRef.current = request;
    // 没有 IntersectionObserver（老内核 / 测试环境）时退化成立即加载
    if (!host || typeof IntersectionObserver === "undefined") {
      request();
      return () => {
        cancelled = true;
      };
    }
    // **懒加载**：进入视口（外扩 200px 预取）才向后端要 data URL。
    // 后端每次调用都要重新解密 + 解包 + base64（实测 11.7ms/张），不能让整篇文档的图一起发请求。
    // root 必须是**文档滚动容器**：用 viewport 当 root 时，容器 overflow 会把目标裁掉，
    // 200px 预取范围等于不存在（滚动到跟前才开始取图，用户会看到"图片加载中…"）。
    const scrollRoot = renderCtx.scrollRootRef.current;
    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          /**
           * **不在这里 `disconnect()`**：加载成功后 `requestedRef` 会一直为 true（自然不再发请求），
           * 而失败/超时后它被清掉 —— 保留观测器才能在"再次进入视口"时重试
           * （用户实测要求：失败后不能永久卡死）。
           */
          request();
        }
      },
      { root: scrollRoot ?? null, rootMargin: "200px" },
    );
    observer.observe(host);
    return () => {
      cancelled = true;
      if (requestRef.current === request) requestRef.current = null;
      observer.disconnect();
    };
  }, [block.media, cacheKey, kind]);

  const caption = block.alt && block.alt.trim() ? block.alt.trim() : null;
  const altText = block.alt ?? block.name ?? "";

  let content: ReactNode;
  if (kind === "external") {
    content = (
      <ImagePlaceholder
        block={block}
        message={`外部链接图片（未保存在文档里，无法离线显示）\n${block.media}`}
      />
    );
  } else if (kind === "vector") {
    content = (
      <ImagePlaceholder
        block={block}
        message={`矢量图（${mediaExtension(block.media) || "EMF/WMF"}），本视图显示不了，请用「用其它应用编辑」在 Word / WPS 里查看`}
      />
    );
  } else if (kind === "unknown") {
    content = (
      <ImagePlaceholder
        block={block}
        message={`暂不支持的图片格式（.${mediaExtension(block.media)}），请用 Word / WPS 查看`}
      />
    );
  } else if (state.status === "ready") {
    /**
     * 显示盒子由 `resolveImageBox` 统一算（与高度估算**同一个函数**）：
     * 图片比版心宽时按比例同时缩宽缩高，拿不到高度时两边都用兜底高度 ——
     * 这样"估算高度 = 渲染高度"恒成立，下一块不会被压在图上。
     */
    const box = resolveImageBox(block, ctx.scale, contentWidth);
    content = (
      <img
        data-docx-img="true"
        src={state.src}
        alt={altText}
        data-docx-img-box={`${box.width}x${box.height}`}
        style={{
          width: `${box.width}px`,
          height: `${box.height}px`,
          // 版心算错时兜底：宽度被 maxWidth 夹住也只是等比留白，不会把高度算歪
          maxWidth: "100%",
          objectFit: "contain",
          display: "block",
        }}
        draggable={false}
      />
    );
  } else if (state.status === "error") {
    content = (
      <ImagePlaceholder
        block={block}
        message={`图片显示不了：${state.message}`}
        onRetry={() => {
          // 手动重试：清掉节流与"已请求"标记，然后直接再发一次（不依赖观测器回调）
          requestedRef.current = false;
          lastAttemptRef.current = 0;
          const retry = requestRef.current;
          if (retry) retry();
          else setState({ status: "idle" });
        }}
      />
    );
  } else {
    // 还没加载：严格按 resolveImageBox 的同一尺寸占位（不夹 MEDIA_MAX_WIDTH/2），
    // 保证占位态与就绪态 offsetHeight 完全一致，首屏测量不会把图片高度记成 272px 导致压字或跨页震荡
    const box = resolveImageBox(block, ctx.scale, contentWidth);
    content = (
      <div
        data-docx-image-skeleton="true"
        style={{
          width: `${box.width}px`,
          height: `${box.height}px`,
          boxSizing: "border-box",
          maxWidth: "100%",
          border: PLACEHOLDER_BORDER,
          borderRadius: "6px",
          background: PLACEHOLDER_BG,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: "11px",
          color: HINT_COLOR,
        }}
      >
        {state.status === "loading" ? "图片加载中…" : "图片（滚动到可见处再加载）"}
      </div>
    );
  }

  return (
    <div
      data-docx-image="true"
      data-docx-media={block.media}
      data-docx-media-kind={kind}
      ref={hostRef}
      style={{ paddingTop: `${BLOCK_MARGIN_Y.image}px`, paddingBottom: `${BLOCK_MARGIN_Y.image}px` }}
    >
      <figure style={{ margin: 0, maxWidth: "100%" }}>
        {content}
        {caption ? (
          <figcaption style={{ marginTop: "3px", fontSize: "11px", color: HINT_COLOR }}>
            {caption}
          </figcaption>
        ) : null}
      </figure>
    </div>
  );
});

/* -------------------------------- 形状 -------------------------------- */

/**
 * 形状块（合同/表单里的线框：下划线、竖线、方框、椭圆）。
 *
 * **不要用 `<hr>`**：合同里的线有精确的起止位置与长度，`<hr>` 撑满整行必然走样。
 * 这里用绝对定位的 `<div>`：
 *  · 水平线 = `border-top`（`height: 0`）；
 *  · 竖线（`vertical`）= `border-left`（`width: 0`）；
 *  · rect / roundRect / ellipse = 带边框（+ 可选填充）的框。
 * 块的盒高 = `yPt + heightPt`（线按线宽兜底），由 `estimateShapeHeight` 参与布局，
 * 所以后面的块不会叠上来。
 */
const DocxShapeBlock = memo(function DocxShapeBlock({
  shape,
  ctx,
}: {
  shape: DocShape;
  ctx: DocxRenderContext;
}) {
  const scale = ctx.scale;
  const lineWidthPt = shape.lineWidthPt !== null && shape.lineWidthPt > 0 ? shape.lineWidthPt : DEFAULT_LINE_WIDTH_PT;
  const lineWidthPx = Math.max(0.5, round2(ptToPx(lineWidthPt, scale)));
  // 颜色缺省用 currentColor：与正文同色，浅色/深色主题下都看得见（Word 的 auto 就是文字色）
  const lineColor = hexColor(shape.lineColor) ?? "currentColor";
  const fill = hexColor(shape.fillColor);
  const borderStyle = shape.dash && shape.dash.trim() ? borderStyleToCss(shape.dash) : "solid";
  const isLine = shape.shape === "line";
  const widthPx = round2(ptToPx(Math.max(0, shape.widthPt), scale));
  const heightPx = round2(ptToPx(Math.max(0, shape.heightPt), scale));
  const left = round2(ptToPx(Math.max(0, shape.xPt), scale));
  const top = round2(ptToPx(Math.max(0, shape.yPt), scale));

  let style: CSSProperties;
  if (isLine && shape.vertical) {
    style = {
      position: "absolute",
      left: `${left}px`,
      top: `${top}px`,
      width: "0px",
      height: `${Math.max(heightPx, lineWidthPx)}px`,
      borderLeft: `${lineWidthPx}px ${borderStyle} ${lineColor}`,
    };
  } else if (isLine) {
    style = {
      position: "absolute",
      left: `${left}px`,
      top: `${top}px`,
      width: `${Math.max(widthPx, 1)}px`,
      height: "0px",
      borderTop: `${lineWidthPx}px ${borderStyle} ${lineColor}`,
    };
  } else {
    style = {
      position: "absolute",
      left: `${left}px`,
      top: `${top}px`,
      width: `${Math.max(widthPx, 1)}px`,
      height: `${Math.max(heightPx, lineWidthPx)}px`,
      border: `${lineWidthPx}px ${borderStyle} ${lineColor}`,
      // 用 backgroundColor 而不是 background 简写：简写在部分内核/测试环境里读不回具体属性
      backgroundColor: fill ?? undefined,
      boxSizing: "border-box",
      borderRadius:
        shape.shape === "ellipse" ? "50%" : shape.shape === "roundRect" ? `${round2(4 * scale)}px` : undefined,
    };
  }

  return (
    <div
      data-docx-shape="true"
      data-docx-shape-kind={shape.shape}
      data-docx-shape-vertical={shape.vertical ? "true" : undefined}
      style={{ position: "relative", width: "100%", height: `${estimateShapeHeight(shape, scale)}px` }}
    >
      <div data-docx-shape-body="true" style={style} />
    </div>
  );
});

/* ------------------------------- 文本框 ------------------------------- */

/**
 * 文本框（`w:txbxContent`）：企业标准的封面整页几乎都是文本框。
 *
 * ── 渲染策略（关系到会不会重叠，重要）────────────────────────────────
 * **默认按文档流渲染**（当作一个普通块：宽度 = `widthPt`、最小高度 = `heightPt`，
 * 画填充/边框，内部块用同一套渲染递归）。理由：文本框在后端是**独立块**，
 * 如果一律绝对定位，`xPt/yPt` 为 0 的那些会全部叠在页顶 —— 比占位卡片还糟。
 * 按流式渲染能保证"内容按顺序完整可读"，而且与高度估算天然同一个口径。
 * **只有"浮动"（`wrap !== "none"`）且 `yPt` 明显非 0（> 12pt）** 时才绝对定位 ——
 * 那才是真的浮在文字上面的对象；其余情况宁可顺序排下去。
 *
 * `height = max(heightPt, 内部内容高 + 内边距×2)`：**内容绝不裁切**（宁可撑高），
 * 与 `resolveTextBoxBox` 共用一套算法，估算高度 = 实测高度。
 */
const DocxTextBoxBlock = memo(function DocxTextBoxBlock({
  box,
  ctx,
  depth,
  contentWidth,
  renderBlocks,
}: {
  box: DocTextBox;
  ctx: DocxRenderContext;
  depth: number;
  contentWidth: number;
  renderBlocks: (blocks: DocBlock[], depth: number, contentWidth: number) => ReactNode;
}) {
  const scale = ctx.scale;
  /**
   * 宽度上限用**纸张可容纳宽**（`ctx.maxBlockWidth`），不是正文版心 ——
   * Word 里文本框是浮动对象，允许超出正文版心。企业标准的日期框声明 482pt，
   * 比版心宽 19px，夹到版心会把「2026-10-21实施」从中间挤断成两行。
   * 左边缘仍对齐版心左边缘，超出的部分向右溢出。
   */
  const maxBoxWidth = ctx.maxBlockWidth > 0 ? ctx.maxBlockWidth : contentWidth;
  const declaredWidth = ptToPx(Math.max(0, box.widthPt), scale);
  /**
   * **布局宽**：夹到版心 —— 文本框不参与撑宽纸张（否则会出现横向滚动条）。
   * **内容宽**：夹到纸张可容纳宽 —— 内容允许视觉溢出到版心外（封面日期框就是靠这个
   * 才能把「2026-10-21发布 … 2026-10-21实施」排成一行，右端 736px 仍在 794px 的纸内）。
   * 估算用的是内容宽，所以"估算高度 = 渲染高度"依然成立。
   */
  const layoutWidth = round2(Math.min(declaredWidth > 0 ? declaredWidth : contentWidth, contentWidth));
  const contentBoxWidth = round2(
    Math.min(declaredWidth > 0 ? declaredWidth : maxBoxWidth, maxBoxWidth),
  );
  const innerWidth = round2(Math.max(24, contentBoxWidth - TEXT_BOX_PADDING * 2));
  const declaredHeight = ptToPx(Math.max(0, box.heightPt), scale);
  const fill = hexColor(box.fillColor);
  const borderColor = hexColor(box.borderColor);
  const borderWidth = Math.max(
    0.5,
    round2(ptToPx(box.borderWidthPt !== null && box.borderWidthPt > 0 ? box.borderWidthPt : TEXT_BOX_DEFAULT_BORDER_PT, scale)),
  );
  /** 只有浮动 + 有明显 y 偏移时才脱离文档流 */
  const floating = box.wrap !== "none" && ptToPx(Math.max(0, box.yPt), scale) > 12;

  return (
    <div
      data-docx-textbox="true"
      data-docx-textbox-wrap={box.wrap}
      data-docx-textbox-floating={floating ? "true" : undefined}
      style={{
        position: floating ? "absolute" : "relative",
        left: floating ? `${round2(ptToPx(Math.max(0, box.xPt), scale))}px` : undefined,
        top: floating ? `${round2(ptToPx(Math.max(0, box.yPt), scale))}px` : undefined,
        // 布局宽夹在版心内：**不撑宽纸张、不出横向滚动条**
        width: `${layoutWidth}px`,
        maxWidth: "100%",
        minHeight: `${round2(declaredHeight)}px`,
        padding: `${TEXT_BOX_PADDING}px`,
        boxSizing: "border-box",
        backgroundColor: fill ?? undefined,
        border: borderColor ? `${borderWidth}px solid ${borderColor}` : undefined,
        borderRadius: floating ? `${round2(3 * scale)}px` : undefined,
      }}
    >
      {/**
       * 内容层：宽度按**纸张可容纳宽**（可以比外框宽），`overflow: visible` —— 内容完整可见、
       * 绝不裁切；外层纸张有 `overflow: clip`，所以内容不会把纸张撑宽。
       */}
      <div
        data-docx-textbox-content="true"
        style={{ width: `${contentBoxWidth}px`, maxWidth: "none", overflow: "visible" }}
      >
        {renderBlocks(box.blocks, depth, innerWidth)}
      </div>
    </div>
  );
});

/* ------------------------------ 不支持的对象 ------------------------------ */

/**
 * 占位卡片：**绝不静默丢失**（OMML 公式、OLE、SmartArt、文本框…）。
 *
 * 连续多个**同类**占位块会被 `planBlocks` 合并成一张卡片（`details.length > 1`）：
 * 标题写成「N 个图形对象（暂不支持显示）」，下面小字逐条列出各自的 `detail` ——
 * 一份合同里 6 张卡片那种视觉噪音就此消掉，内容仍然查得到（复制文本里也仍有 `[label]`）。
 */
export function UnsupportedCard({
  label,
  details,
}: {
  label: string;
  /** 组内每个块的 detail（单个占位块就是长度 1 的数组） */
  details: readonly string[];
}) {
  const count = details.length;
  // 单个：文案与样式保持原样；合并后：「N 个图形对象（暂不支持显示）」
  const title = count > 1 ? `${count} 个${label}` : label;
  return (
    <div
      data-docx-unsupported="true"
      data-docx-label={label}
      data-docx-unsupported-count={count}
      style={{ paddingTop: `${BLOCK_MARGIN_Y.unsupported}px`, paddingBottom: `${BLOCK_MARGIN_Y.unsupported}px` }}
    >
      <div
        style={{
          border: PLACEHOLDER_BORDER,
          borderRadius: "6px",
          background: PLACEHOLDER_BG,
          padding: "5px 10px",
        }}
      >
        <div style={{ fontSize: "12px", color: HINT_COLOR }}>{title}</div>
        {details.map((detail, index) =>
          detail ? (
            <div
              key={index}
              style={{
                marginTop: "2px",
                fontSize: "11px",
                lineHeight: "16px",
                color: HINT_COLOR,
                whiteSpace: "pre-wrap",
                wordBreak: "break-word",
              }}
            >
              {/* 合并卡片里逐条编号，方便和原文对照 */}
              {count > 1 ? `${index + 1}. ${detail}` : detail}
            </div>
          ) : null,
        )}
      </div>
    </div>
  );
}

/**
 * 「图形对象（暂不支持显示）」这类 label 直接接数字前缀即可：
 * 合并后读作「6 个图形对象（暂不支持显示）」。
 */

/* -------------------------------- 分派 -------------------------------- */

function renderBlocksInFlow(
  blocks: DocBlock[],
  ctx: DocxRenderContext,
  depth: number,
  contentWidth: number,
): ReactNode {
  if (blocks.length === 0) return null;
  return blocks.map((block, index) => (
    <DocxBlock key={index} block={block} ctx={ctx} depth={depth} contentWidth={contentWidth} />
  ));
}

export const DocxBlock = memo(function DocxBlock({
  block,
  ctx,
  depth = 0,
  contentWidth,
}: DocxBlockProps) {
  switch (block.kind) {
    case "paragraph":
      return <DocxParagraph paragraph={block} ctx={ctx} />;
    case "table":
      return (
        <DocxTable
          table={block}
          ctx={ctx}
          depth={depth}
          contentWidth={contentWidth}
          renderBlocks={(blocks, nextDepth, cellWidth) =>
            renderBlocksInFlow(blocks, ctx, nextDepth, cellWidth)
          }
        />
      );
    case "image":
      return <DocxImageBlock block={block} ctx={ctx} contentWidth={contentWidth} />;
    case "shape":
      return <DocxShapeBlock shape={block} ctx={ctx} />;
    case "textBox":
      return (
        <DocxTextBoxBlock
          box={block}
          ctx={ctx}
          depth={depth + 1}
          contentWidth={contentWidth}
          renderBlocks={(blocks, nextDepth, innerWidth) =>
            renderBlocksInFlow(blocks, ctx, nextDepth, innerWidth)
          }
        />
      );
    case "pageBreak":
      return <PageBreakLine label="分页符" />;
    case "unsupported":
      return <UnsupportedCard label={block.label} details={[block.detail]} />;
    default:
      return null;
  }
});
