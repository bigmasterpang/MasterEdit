import { useEffect, useRef, useState } from "react";
import { convertFileSrc } from "@tauri-apps/api/core";
import { loadKatex, loadMermaid } from "../../utils/markdown";
import { isExternalUrl, resolveLocalImagePath, joinPath, dirName } from "../../utils/filePath";
import { openExternal, openPath } from "../../utils/fileActions";
import { toggleTaskOnLine } from "../../utils/editorCommands";
import { sanitizeHtml } from "../../utils/sanitize";
import { useAppStore } from "../../stores/appStore";
import { useSearchStore } from "../../stores/searchStore";
import { buildRegex, type SearchOptions } from "../../utils/searchEngine";
import { ContextMenu } from "../common/ContextMenu";

interface Props {
  /** 本预览所属文档：解析相对图片/链接必须用它，不能用全局激活文档 */
  docId?: string | null;
  html: string;
  hasMath: boolean;
  hasMermaid: boolean;
  isDark: boolean;
  scrollRef: React.RefObject<HTMLDivElement | null>;
}

let mermaidSeq = 0;

/**
 * Markdown 预览。
 * 顺序：清理 HTML -> 注入 DOM -> 修正图片路径/链接 -> 延迟渲染 KaTeX / Mermaid。
 */
export function MarkdownPreview({
  docId,
  html,
  hasMath,
  hasMermaid,
  isDark,
  scrollRef,
}: Props) {
  const bodyRef = useRef<HTMLDivElement>(null);
  const generationRef = useRef(0);
  const [markCount, setMarkCount] = useState(0);
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);

  const searchVisible = useSearchStore((s) => s.visible);
  const searchQuery = useSearchStore((s) => s.query);
  const searchOptions = useSearchStore((s) => s.options);
  const searchCurrent = useSearchStore((s) => s.current);

  useEffect(() => {
    const root = bodyRef.current;
    if (!root) return;
    const generation = ++generationRef.current;

    const safe = sanitizeHtml(html);
    root.innerHTML = safe;

    // 必须按「本预览所属文档」解析相对路径：用全局 activeId 的话，
    // 双栏打开两个不同目录的文档时，非激活栏的图片会按另一篇文档的目录去解析
    const doc = docId
      ? useAppStore.getState().docs.find((d) => d.id === docId) ?? null
      : null;
    const docPath = doc?.filePath ?? null;

    /* ---------------- 图片路径修正 ---------------- */
    root.querySelectorAll<HTMLImageElement>("img").forEach((img) => {
      const src = img.getAttribute("src") ?? "";
      if (!src) return;
      if (/^(https?:|data:|asset:)/i.test(src)) return;

      const localPath = resolveLocalImagePath(docPath, src);
      if (!localPath) {
        markBroken(img, "文档尚未保存，无法解析相对路径图片");
        return;
      }
      img.setAttribute("data-local-path", localPath);
      img.setAttribute("src", convertFileSrc(localPath));
      img.addEventListener("error", () => markBroken(img, "图片无法加载"), {
        once: true,
      });
    });

    /* ---------------- 链接拦截 ---------------- */
    root.querySelectorAll<HTMLAnchorElement>("a[href]").forEach((anchor) => {
      anchor.addEventListener("click", (event) => {
        const href = anchor.getAttribute("href") ?? "";
        if (!href || href.startsWith("#")) {
          event.preventDefault();
          if (href.startsWith("#")) {
            const target = root.querySelector<HTMLElement>(
              `[id="${CSS.escape(href.slice(1))}"]`,
            );
            target?.scrollIntoView({ behavior: "smooth", block: "start" });
          }
          return;
        }
        event.preventDefault();
        if (isExternalUrl(href)) {
          if (href.startsWith("#")) return;
          void openExternal(href);
          return;
        }
        // 相对路径：如果是 Markdown 文件则在应用内打开
        if (docPath && /\.(md|markdown|mdown|mkd|mkdn|txt)$/i.test(href)) {
          const base = href.split("#")[0].split("?")[0];
          let decoded = base;
          try {
            decoded = decodeURIComponent(base);
          } catch {
            /* 保留原值 */
          }
          void openPath(joinPath(dirName(docPath), decoded));
          return;
        }
        if (docPath) {
          void openExternal(convertFileSrc(joinPath(dirName(docPath), href)));
        }
      });
    });

    /* ---------------- 任务列表：预览内可勾选，回写源码 ---------------- */
    root.querySelectorAll<HTMLInputElement>("input.task-box").forEach((box) => {
      box.disabled = false;
      box.removeAttribute("disabled");
      box.addEventListener("change", () => {
        const line = Number(box.getAttribute("data-line"));
        if (Number.isFinite(line)) toggleTaskOnLine(line, docId ?? undefined);
        else box.checked = !box.checked;
      });
    });

    /* ---------------- KaTeX（按需加载） ---------------- */
    if (hasMath) {
      void loadKatex().then((katex) => {
        if (generationRef.current !== generation) return;
        root.querySelectorAll<HTMLElement>(".math-inline, .math-block").forEach((el) => {
          const tex = el.dataset.tex ?? "";
          const displayMode = el.classList.contains("math-block");
          try {
            el.innerHTML = katex.renderToString(tex, {
              displayMode,
              throwOnError: false,
              strict: "ignore",
              output: "html",
            });
          } catch (error) {
            el.classList.add("math-error");
            el.textContent = `公式解析失败: ${tex} (${String(error)})`;
          }
        });
      });
    }

    /* ---------------- Mermaid（按需加载） ---------------- */
    if (hasMermaid) {
      void loadMermaid(isDark).then(async (mermaid) => {
        if (generationRef.current !== generation) return;
        const blocks = Array.from(
          root.querySelectorAll<HTMLElement>("pre > code.language-mermaid"),
        );
        for (const code of blocks) {
          if (generationRef.current !== generation) return;
          const pre = code.parentElement;
          const graph = code.textContent ?? "";
          const target = document.createElement("div");
          target.className = "mermaid-block";
          try {
            const id = `masteredit-mermaid-${++mermaidSeq}`;
            const { svg } = await mermaid.render(id, graph);
            target.innerHTML = svg;
          } catch (error) {
            target.classList.add("mermaid-error");
            target.textContent = `Mermaid 渲染失败：${String(error)}`;
          }
          pre?.replaceWith(target);
        }
      });
    }
  }, [docId, html, hasMath, hasMermaid, isDark]);

  /* ---------------- 搜索高亮（作用于渲染后的文本） ---------------- */
  useEffect(() => {
    const root = bodyRef.current;
    if (!root) return;
    clearSearchMarks(root);
    if (!searchVisible || !searchQuery.trim()) {
      setMarkCount(0);
      return;
    }
    setMarkCount(highlightText(root, searchQuery, searchOptions));
  }, [html, searchVisible, searchQuery, searchOptions]);

  useEffect(() => {
    const root = bodyRef.current;
    if (!root || markCount === 0) return;
    const marks = root.querySelectorAll<HTMLElement>("mark.search-hit");
    if (marks.length === 0) return;
    const index = Math.min(Math.max(0, searchCurrent), marks.length - 1);
    marks.forEach((mark, i) => mark.classList.toggle("current", i === index));
    marks[index]?.scrollIntoView({ block: "center", behavior: "smooth" });
  }, [searchCurrent, markCount]);

  /* ---------------- 右键菜单（预览区无法编辑，仅提供复制类操作） ---------------- */
  useEffect(() => {
    const root = scrollRef.current;
    if (!root) return;
    const onContextMenu = (event: MouseEvent) => {
      event.preventDefault();
      setMenu({ x: event.clientX, y: event.clientY });
    };
    root.addEventListener("contextmenu", onContextMenu);
    return () => root.removeEventListener("contextmenu", onContextMenu);
  }, [scrollRef]);

  return (
    <>
      <div
        ref={scrollRef}
        className="print-plain h-full overflow-auto"
        style={{ scrollBehavior: "auto" }}
      >
        <div className="print-content mx-auto w-full px-8 py-6" style={{ maxWidth: 900 }}>
          <div ref={bodyRef} className="md-body" />
        </div>
      </div>
      {menu ? (
        <ContextMenu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          groups={[
            [
              {
                label: "复制选中文本",
                hint: "Ctrl+C",
                onClick: () => {
                  const text = window.getSelection()?.toString() ?? "";
                  if (text) void navigator.clipboard.writeText(text);
                },
              },
              {
                label: "复制全文（纯文本）",
                onClick: () => {
                  const text = scrollRef.current?.innerText ?? "";
                  if (text) void navigator.clipboard.writeText(text);
                },
              },
              {
                label: "全选",
                hint: "Ctrl+A",
                onClick: () => {
                  const range = document.createRange();
                  if (!scrollRef.current) return;
                  range.selectNodeContents(scrollRef.current);
                  const selection = window.getSelection();
                  selection?.removeAllRanges();
                  selection?.addRange(range);
                },
              },
            ],
          ]}
        />
      ) : null}
    </>
  );
}

/** 清除预览中的搜索高亮 */
function clearSearchMarks(root: HTMLElement): void {
  root.querySelectorAll("mark.search-hit").forEach((mark) => {
    const parent = mark.parentNode;
    if (!parent) return;
    while (mark.firstChild) parent.insertBefore(mark.firstChild, mark);
    parent.removeChild(mark);
    parent.normalize();
  });
}

/** 在渲染后的文本节点中高亮匹配，返回匹配数量（用于预览模式的定位） */
function highlightText(root: HTMLElement, query: string, options: SearchOptions): number {
  const { regex } = buildRegex(query, options);
  if (!regex) return 0;

  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode: (node) => {
      const parent = node.parentElement;
      if (!parent) return NodeFilter.FILTER_REJECT;
      const tag = parent.tagName;
      if (tag === "SCRIPT" || tag === "STYLE") return NodeFilter.FILTER_REJECT;
      return node.nodeValue && node.nodeValue.trim()
        ? NodeFilter.FILTER_ACCEPT
        : NodeFilter.FILTER_REJECT;
    },
  });
  const targets: Text[] = [];
  let node = walker.nextNode();
  while (node) {
    targets.push(node as Text);
    node = walker.nextNode();
  }

  let count = 0;
  for (const textNode of targets) {
    const text = textNode.nodeValue ?? "";
    regex.lastIndex = 0;
    let match: RegExpExecArray | null;
    let matched = false;
    let cursor = 0;
    const fragment = document.createDocumentFragment();
    while ((match = regex.exec(text)) !== null) {
      if (match[0].length === 0) {
        regex.lastIndex += 1;
        continue;
      }
      matched = true;
      if (match.index > cursor) {
        fragment.appendChild(document.createTextNode(text.slice(cursor, match.index)));
      }
      const mark = document.createElement("mark");
      mark.className = count === 0 ? "search-hit current" : "search-hit";
      mark.textContent = match[0];
      fragment.appendChild(mark);
      count += 1;
      cursor = match.index + match[0].length;
      if (regex.lastIndex <= match.index) regex.lastIndex = match.index + 1;
    }
    if (!matched) continue;
    if (cursor < text.length) {
      fragment.appendChild(document.createTextNode(text.slice(cursor)));
    }
    textNode.parentNode?.replaceChild(fragment, textNode);
  }
  return count;
}

function markBroken(img: HTMLImageElement, reason: string) {
  img.classList.add("md-img-broken");
  img.removeAttribute("src");
  img.setAttribute("alt", `${img.getAttribute("alt") ?? ""}（${reason}）`);
  img.title = reason;
}
