/**
 * 大纲侧栏：按 `outlineLevel`（0..8）收集标题，点击跳转到对应块。
 *
 * **只看 `outlineLevel`**，不看 `styleId` / `style`：Word、WPS 与各类模板里的样式 id
 * 五花八门（`4`、`a9`、`heading 4`…），用样式名判断标题必然漏掉一批；
 * 后端已经把「样式继承链」算成了大纲级别，这里直接用。
 */
import { useState } from "react";
import { Icon } from "../common/Icon";
import type { OutlineItem } from "./docxGrid";

export interface DocxOutlineProps {
  items: OutlineItem[];
  /** 还在后台取块（大纲可能不全） */
  loading: boolean;
  /** 是否只取到了文档的一部分块 */
  partial: boolean;
  onJump(index: number): void;
}

/** 每一级标题的缩进（px） */
const LEVEL_INDENT = 13;

export function DocxOutline({ items, loading, partial, onJump }: DocxOutlineProps) {
  const [selected, setSelected] = useState<number | null>(null);

  return (
    <aside
      data-docx-outline="true"
      data-docx-outline-count={items.length}
      className="flex w-[236px] shrink-0 flex-col border-r border-line bg-panel"
    >
      <div className="flex items-center gap-1.5 border-b border-line px-3 py-2 text-[12px] text-muted">
        <Icon name="list" size={13} />
        <span className="font-medium text-fg">文档大纲</span>
        <span className="text-faint">
          {items.length > 0 ? `${items.length} 个标题` : loading ? "读取中…" : "无标题"}
        </span>
      </div>
      <div className="min-h-0 flex-1 overflow-auto py-1">
        {items.length === 0 ? (
          <div className="px-3 py-2 text-[11px] leading-relaxed text-faint">
            {loading
              ? "正在读取文档内容…"
              : "这篇文档没有设置大纲级别（标题样式）的段落。"}
          </div>
        ) : (
          items.map((item, order) => (
            <button
              key={`${item.index}-${order}`}
              type="button"
              data-docx-outline-item="true"
              data-docx-outline-level={item.level}
              data-docx-outline-index={item.index}
              title={`${item.text}\n（第 ${item.index + 1} 块，级别 ${item.level + 1}）`}
              onClick={() => {
                setSelected(order);
                onJump(item.index);
              }}
              className={`flex w-full items-baseline gap-1.5 px-2 py-[3px] text-left text-[12px] leading-[1.5] hover:bg-hover ${
                selected === order ? "bg-accent/10 text-accent" : "text-fg/90"
              } ${item.level === 0 ? "font-medium" : ""}`}
              style={{ paddingLeft: `${8 + item.level * LEVEL_INDENT}px` }}
            >
              <span className="text-[10px] text-faint">H{item.level + 1}</span>
              <span className="min-w-0 flex-1 truncate">{item.text || "（空标题）"}</span>
            </button>
          ))
        )}
      </div>
      {partial ? (
        <div className="border-t border-line px-3 py-2 text-[11px] leading-relaxed text-faint">
          文档较大，大纲只覆盖已读取的部分内容。
        </div>
      ) : null}
    </aside>
  );
}
