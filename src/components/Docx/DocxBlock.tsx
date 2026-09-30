/**
 * 单个块的渲染：段落（含 run 级格式、列表前缀、四边边框）、图片（懒加载）、
 * 形状（线框）、分页提示线、占位卡片。
 *
 * 只读原则（见 `docs/plan-docx.md`）：**不做 contentEditable、不做输入框**，
 * 文字保持可选中复制（不加 `user-select: none`）。
 */
import { memo, useEffect, useRef, useState, type CSSProperties, type ReactNode } from "react";
import type { DocBlock, DocImage, DocParagraph, DocShape } from "../../types";
import {
  BLOCK_MARGIN_Y,
  DEFAULT_LINE_WIDTH_PT,
  PAGE_BREAK_HEIGHT,
  borderStyleToCss,
  breakHintText,
  estimateShapeHeight,
  hexColor,
  listGapPx,
  paragraphBoxStyle,
  ptToPx,
  round2,
  runInlineStyle,
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
    children.push(
      <span key={`run-${index}`} data-docx-run="true" style={runInlineStyle(run, ctx.scale)}>
        {run.text}
      </span>,
    );
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
function ImagePlaceholder({ block, message }: { block: DocImage; message: string }) {
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
    </div>
  );
}

const DocxImageBlock = memo(function DocxImageBlock({
  block,
  ctx,
}: {
  block: DocImage;
  ctx: DocxRenderContext;
}) {
  const kind = mediaKind(block.media);
  const [state, setState] = useState<ImageState>({ status: "idle" });
  const hostRef = useRef<HTMLDivElement | null>(null);
  /** 同一个图片块只请求一次（重复渲染 / 视口反复进出都不会重发） */
  const requestedRef = useRef(false);
  const cacheKey = ctx.mediaCacheKey(block.media);

  useEffect(() => {
    if (kind !== "raster") return;
    // 命中缓存（同一张图在前面已经取过）就直接用，不再走 invoke
    const cached = ctx.mediaCache.get(cacheKey);
    if (cached) {
      setState({ status: "ready", src: cached });
      return;
    }
    let cancelled = false;
    const request = () => {
      if (requestedRef.current) return;
      requestedRef.current = true;
      setState({ status: "loading" });
      ctx
        .loadMedia(block.media)
        .then((src) => {
          if (cancelled) return;
          setState({ status: "ready", src });
        })
        .catch((reason: unknown) => {
          if (cancelled) return;
          const message = typeof reason === "string" && reason.trim() ? reason : "图片读取失败";
          setState({ status: "error", message });
        });
    };

    const host = hostRef.current;
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
    const scrollRoot = ctx.scrollRootRef.current;
    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          observer.disconnect();
          request();
        }
      },
      { root: scrollRoot ?? null, rootMargin: "200px" },
    );
    observer.observe(host);
    return () => {
      cancelled = true;
      observer.disconnect();
    };
  }, [block.media, cacheKey, ctx, kind]);

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
    const width = block.widthPx > 0 ? `${round2(block.widthPx * ctx.scale)}px` : undefined;
    const height = block.heightPx > 0 ? `${round2(block.heightPx * ctx.scale)}px` : undefined;
    content = (
      <img
        data-docx-img="true"
        src={state.src}
        alt={altText}
        /**
         * 按文档给的 widthPx/heightPx 显示（模型的尺寸就是 Word 里的显示尺寸）；
         * `object-fit: contain` 保证文档里非等比的尺寸也不会把图拉变形（多出来的部分留白）。
         */
        style={{ width, height, maxWidth: "100%", objectFit: "contain", display: "block" }}
        draggable={false}
      />
    );
  } else if (state.status === "error") {
    content = <ImagePlaceholder block={block} message={`图片显示不了：${state.message}`} />;
  } else {
    // 还没加载：按文档给的尺寸占位（滚动时高度不会跳变），不请求后端
    const width = block.widthPx > 0 ? Math.min(block.widthPx, MEDIA_MAX_WIDTH) : 220;
    const height = block.heightPx > 0 ? Math.min(block.heightPx, 200) : 56;
    content = (
      <div
        data-docx-image-skeleton="true"
        style={{
          width: `${round2(width * ctx.scale)}px`,
          height: `${round2(height * ctx.scale)}px`,
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

function renderBlocksInFlow(blocks: DocBlock[], ctx: DocxRenderContext, depth: number): ReactNode {
  if (blocks.length === 0) return null;
  return blocks.map((block, index) => (
    <DocxBlock key={index} block={block} ctx={ctx} depth={depth} />
  ));
}

export const DocxBlock = memo(function DocxBlock({ block, ctx, depth = 0 }: DocxBlockProps) {
  switch (block.kind) {
    case "paragraph":
      return <DocxParagraph paragraph={block} ctx={ctx} />;
    case "table":
      return (
        <DocxTable
          table={block}
          ctx={ctx}
          depth={depth}
          renderBlocks={(blocks, nextDepth) => renderBlocksInFlow(blocks, ctx, nextDepth)}
        />
      );
    case "image":
      return <DocxImageBlock block={block} ctx={ctx} />;
    case "shape":
      return <DocxShapeBlock shape={block} ctx={ctx} />;
    case "pageBreak":
      return <PageBreakLine label="分页符" />;
    case "unsupported":
      return <UnsupportedCard label={block.label} details={[block.detail]} />;
    default:
      return null;
  }
});
