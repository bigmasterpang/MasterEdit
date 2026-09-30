/**
 * DOCX 渲染的共享上下文与媒体分类（纯逻辑，避免 DocxBlock ↔ DocxTable 互相 import 成环）。
 */
import type { DocTable } from "../../types";

/** 渲染上下文：整篇文档共用一份（字号倍率、图片缓存、查找高亮、表格右键菜单） */
export interface DocxRenderContext {
  /** 字号/尺寸缩放倍率（Ctrl+滚轮） */
  scale: number;
  /**
   * 图片 data URL 缓存：**key 必须是 `路径|media`**（见 `mediaCacheKey`）。
   * 同一个 media 只请求一次 —— 后端每次 `document_media` 都要内存解密 + 解包 + base64，
   * 实测约 11.7ms；且每次返回的都是新字符串，重复塞 `<img src>` 会重新解码。
   */
  mediaCache: Map<string, string>;
  /** 缓存 key：`文档路径|包内路径`（同一个 media 在不同文档里不是同一张图） */
  mediaCacheKey(media: string): string;
  /** 取图片 data URL（内部会去重 + 写回 mediaCache；不进视口的图片不会被调用） */
  loadMedia(media: string): Promise<string>;
  /** 当前查找高亮的块下标（null = 没有） */
  highlightBlock: number | null;
  /**
   * 该块**所在页的页码**（1 起）与总页数：`PAGE` / `NUMPAGES` 域用它替换缓存文本。
   * 正文由 `DocxBlocks` 逐页传入（同一页的块共用一个 ctx）；页眉页脚按所在页传。
   */
  pageNumber: number;
  totalPages: number;
  /**
   * **纸张可容纳块宽**（px）：文本框这类浮动对象允许超出正文版心、只受纸张边界约束，
   * 用它当夹取上限（左边缘仍对齐版心左边缘，超出部分向右溢出）。
   * 普通块不用它（仍按版心 / 单元格内宽排版）。
   */
  maxBlockWidth: number;
  /**
   * 滚动容器（虚拟滚动那个 div）的 ref。图片懒加载的 IntersectionObserver
   * **必须以它为 root**：文档区自己就是滚动容器，若用 viewport 当 root，
   * 目标会被容器的 overflow 裁掉，`rootMargin` 的预取范围完全失效
   * （视口下方那 200px 不算数 → 滚动时才开始取图，会看到"图片加载中…"）。
   * 由 `DocxBlocks` 在 layout 阶段同步进来。
   */
  scrollRootRef: { current: HTMLElement | null };
  /** 纸张底色（默认 `#ffffff` 白纸；用户可在工具栏切换或自定义） */
  pageBg?: string;
  /** 表格右键菜单（复制为制表符文本等）；不注入则不显示菜单 */
  onTableContextMenu?: (event: { clientX: number; clientY: number }, table: DocTable) => void;
}

/** 递归渲染的最大深度：嵌套表格最多 6 层，防病态文档把栈撑爆 */
export const MAX_BLOCK_DEPTH = 6;

export type MediaKind =
  /** 浏览器能直接显示的位图/矢量图（png/jpg/gif/bmp/webp/svg…） */
  | "raster"
  /** EMF / WMF：Word 里粘贴 Excel 图表得到的就是它，`<img>` 显示不出来 */
  | "vector"
  /** 不在包内（`TargetMode="External"` 的 http(s) 链接图片，后端会拒绝） */
  | "external"
  /** 包内其它格式（tif / bin…），浏览器同样显示不了 */
  | "unknown";

const RASTER_EXTENSIONS = new Set(["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg", "ico", "avif"]);
const VECTOR_EXTENSIONS = new Set(["emf", "wmf"]);

/**
 * 判断图片该怎么处理。三条已经确认的约束（后端契约）：
 *  · `word/media/` 之外的 `media` 是外部链接图片 → **不要调 `document_media`**（后端会拒绝）；
 *  · EMF / WMF 能取到 data URL，但浏览器渲染不出来 → 占位卡片，不要塞 `<img>`；
 *  · 其余按位图正常懒加载。
 */
export function mediaKind(media: string): MediaKind {
  const lower = media.toLowerCase();
  if (!lower.startsWith("word/media/")) return "external";
  const dot = lower.lastIndexOf(".");
  const ext = dot >= 0 ? lower.slice(dot + 1) : "";
  if (RASTER_EXTENSIONS.has(ext)) return "raster";
  if (VECTOR_EXTENSIONS.has(ext)) return "vector";
  return "unknown";
}

/** 图片扩展名（占位文案里显示用） */
export function mediaExtension(media: string): string {
  const lower = media.toLowerCase();
  const dot = lower.lastIndexOf(".");
  return dot >= 0 ? lower.slice(dot + 1).toUpperCase() : "";
}
