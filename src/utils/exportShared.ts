import { invoke } from "@tauri-apps/api/core";
import { extName } from "./filePath";

export const MIME_BY_EXT: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
  svg: "image/svg+xml",
};

/** 复制预览节点：去掉搜索高亮标记，得到干净的 HTML */
export function clonePreview(source: HTMLElement): HTMLElement {
  const clone = source.cloneNode(true) as HTMLElement;
  clone.querySelectorAll("mark.search-hit").forEach((mark) => {
    const parent = mark.parentNode;
    if (!parent) return;
    while (mark.firstChild) parent.insertBefore(mark.firstChild, mark);
    parent.removeChild(mark);
    parent.normalize();
  });
  clone.querySelectorAll("input[type=checkbox]").forEach((box) => {
    box.removeAttribute("disabled");
  });
  return clone;
}

/** 把本地图片替换为 Base64 data URI，返回处理数量 */
export async function inlineLocalImages(root: HTMLElement): Promise<number> {
  const images = Array.from(root.querySelectorAll("img"));
  let count = 0;
  for (const img of images) {
    const localPath = img.getAttribute("data-local-path");
    if (!localPath) continue;
    try {
      const base64 = await invoke<string>("read_file_as_base64", {
        path: localPath,
      });
      const mime = MIME_BY_EXT[extName(localPath)] ?? "image/png";
      img.setAttribute("src", `data:${mime};base64,${base64}`);
      img.removeAttribute("data-local-path");
      count += 1;
    } catch {
      /* 单张图片失败不影响整体导出 */
    }
  }
  return count;
}


/** 读取图片二进制与尺寸（用于 DOCX 内嵌图片） */
export async function loadImageBinary(
  dataUrl: string,
): Promise<{ data: Uint8Array; width: number; height: number } | null> {
  try {
    const response = await fetch(dataUrl);
    const buffer = await response.arrayBuffer();
    const size = await new Promise<{ width: number; height: number }>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
      image.onerror = () => reject(new Error("图片解析失败"));
      image.src = dataUrl;
    });
    return { data: new Uint8Array(buffer), width: size.width, height: size.height };
  } catch {
    return null;
  }
}
