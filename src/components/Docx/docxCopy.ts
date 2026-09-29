/**
 * DOCX 视图的复制能力（**只读视图唯一的"输出"通道**，用户明确要求）。
 *
 * 三条出口：
 *  1. 正文文字本身可选中复制（渲染时不加 `user-select: none`，不做 contentEditable）；
 *  2. 「复制全部文本」→ 把所有块的纯文本拼起来（见 docxTable 的 blocksToPlainText）；
 *  3. 表格右键「复制为制表符文本」→ 行内 `\t`、行间 `\n`，可直接粘进 Excel。
 *
 * 剪贴板优先用 `navigator.clipboard.writeText`（WebView2 支持）；
 * 它可能因为权限/非聚焦而 reject，此时退回 `textarea + execCommand("copy")`。
 * 两条都失败时给**中文提示**（不静默失败）。
 */

export interface CopyOutcome {
  ok: boolean;
  /** 给用户看的中文提示（成功也提示，让用户知道复制了多少） */
  message: string;
}

/** textarea + execCommand 兜底：老 WebView 或 clipboard API 被安全策略挡住时用 */
function copyViaTextarea(text: string): boolean {
  if (typeof document === "undefined" || typeof document.execCommand !== "function") return false;
  const area = document.createElement("textarea");
  area.value = text;
  // 放到视口外，避免复制时页面闪一下
  area.setAttribute("readonly", "readonly");
  area.style.position = "fixed";
  area.style.top = "-2000px";
  area.style.left = "-2000px";
  document.body.appendChild(area);
  try {
    area.select();
    return document.execCommand("copy");
  } catch {
    return false;
  } finally {
    area.remove();
  }
}

/** 复制纯文本到剪贴板；返回中文结果提示 */
export async function copyTextToClipboard(text: string): Promise<CopyOutcome> {
  if (!text) return { ok: false, message: "没有可复制的内容" };
  try {
    const clipboard = typeof navigator === "undefined" ? undefined : navigator.clipboard;
    if (clipboard && typeof clipboard.writeText === "function") {
      await clipboard.writeText(text);
      return { ok: true, message: `已复制 ${text.length} 个字符` };
    }
  } catch {
    // 落到下面的兜底路径
  }
  if (copyViaTextarea(text)) return { ok: true, message: `已复制 ${text.length} 个字符` };
  return {
    ok: false,
    message: "复制失败：系统拒绝了剪贴板访问，请手动选中文字后按 Ctrl+C",
  };
}
