import { useState, useRef, useEffect } from "react";
import { invoke } from "@tauri-apps/api/core";
import { openPath as openWithSystem, revealItemInDir } from "@tauri-apps/plugin-opener";
import { Icon, type IconName } from "../common/Icon";
import { ContextMenu, type ContextMenuItem } from "../common/ContextMenu";
import { useAppStore } from "../../stores/appStore";
import { useTabDragStore } from "../../stores/tabDragStore";
import {
  closeAllDocsWithConfirm,
  closeDocWithConfirm,
  closeOtherDocsWithConfirm,
  closeRightDocsWithConfirm,
  newDocument,
  openActiveDocWithExternalApp,
  openFileDialog,
  reloadDocFromDisk,
  saveDoc,
  saveDocAs,
  toggleDocReadOnly,
} from "../../utils/fileActions";
import { getDocBaseName, getDocTitle, isPdfDoc } from "../../utils/filePath";
import type { DocState } from "../../types";

interface Props {
  pane?: 0 | 1;
}

interface ContextMenuState {
  x: number;
  y: number;
  doc: DocState;
}

/** 多标签页栏：按分栏渲染标签 + 右键菜单 + 指针拖拽排序/分栏 + 滚轮切换标签 */
export function TabBar({ pane = 0 }: Props) {
  const docs = useAppStore((s) => s.docs);
  const activeIds = useAppStore((s) => s.activeIds);
  const activeId = useAppStore((s) => s.activeId);
  const layout = useAppStore((s) => s.layout);
  const gameOpen = useAppStore((s) => s.gameOpen);

  const dragStore = useTabDragStore();

  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [plusMenu, setPlusMenu] = useState<{ x: number; y: number } | null>(null);

  const containerRef = useRef<HTMLDivElement>(null);
  const activeTabRef = useRef<HTMLDivElement>(null);
  const lastWheelTimeRef = useRef(0);

  // 本栏中的文档
  const paneDocs = docs.filter((d) => (d.pane ?? 0) === pane);
  const paneActiveId = activeIds[pane] ?? (pane === layout.activePane ? activeId : null);

  // 当激活标签改变时，自动使其进入视口
  useEffect(() => {
    if (activeTabRef.current) {
      activeTabRef.current.scrollIntoView({ inline: "nearest", block: "nearest", behavior: "smooth" });
    }
  }, [paneActiveId]);

  // 滚轮在标签栏滚动切换标签
  const handleWheel = (e: React.WheelEvent) => {
    if (paneDocs.length <= 1) return;
    const delta = Math.abs(e.deltaY) >= Math.abs(e.deltaX) ? e.deltaY : e.deltaX;
    if (Math.abs(delta) < 5) return;

    // 节流 120ms，避免普通滚轮滚动一下切换多格
    const now = Date.now();
    if (now - lastWheelTimeRef.current < 120) return;
    lastWheelTimeRef.current = now;
    e.preventDefault();

    const currentIdx = paneDocs.findIndex((d) => d.id === paneActiveId);
    const safeIdx = currentIdx >= 0 ? currentIdx : 0;
    const nextIdx =
      delta > 0
        ? (safeIdx + 1) % paneDocs.length
        : (safeIdx - 1 + paneDocs.length) % paneDocs.length;
    useAppStore.getState().activateDoc(paneDocs[nextIdx].id, pane);
  };

  // 无文档且未开启江湖时隐藏标签栏
  if (docs.length === 0 && !gameOpen) return null;

  const getContextMenuGroups = (doc: DocState): ContextMenuItem[][] => {
    return [
      // 1. 关闭操作
      [
        {
          label: "关闭",
          hint: "Ctrl+W",
          onClick: () => void closeDocWithConfirm(doc.id),
        },
        {
          label: "关闭其他标签页",
          disabled: paneDocs.length <= 1,
          // 必须走带确认的关闭流程：直接调用 store 的 closeOtherDocs 会丢弃
          // 其它标签页里未保存的内容，且不可恢复
          onClick: () => void closeOtherDocsWithConfirm(doc.id),
        },
        {
          label: "关闭右侧标签页",
          disabled: paneDocs[paneDocs.length - 1]?.id === doc.id,
          onClick: () => void closeRightDocsWithConfirm(doc.id),
        },
        {
          label: "关闭全部标签页",
          onClick: () => void closeAllDocsWithConfirm(),
        },
      ],
      // 2. 磁盘与复制
      [
        ...(doc.filePath
          ? [
              {
                label: "从磁盘重新加载",
                icon: "refresh" as IconName,
                onClick: () => void reloadDocFromDisk(doc.id),
              },
              {
                label: "复制完整路径",
                icon: "copy" as IconName,
                onClick: () => void navigator.clipboard.writeText(doc.filePath!),
              },
            ]
          : []),
        {
          label: "复制文件名",
          icon: "file-text" as IconName,
          onClick: () => void navigator.clipboard.writeText(getDocBaseName(doc)),
        },
      ],
      // 3. 操作系统相关
      ...(doc.filePath
        ? [
            [
              {
                label: doc.readOnly ? "切换为可编辑" : "切换为只读",
                icon: "lock" as IconName,
                onClick: () => void toggleDocReadOnly(doc.id),
              },
              {
                label: "用系统默认程序打开",
                icon: "external-link" as IconName,
                // 表格只读、PDF 排版保真等场景下，交给系统里的 Office / WPS / 阅读器处理
                onClick: () => {
                  void openWithSystem(doc.filePath!).catch(async (error) => {
                    const { showMessage } = await import("../../stores/dialogStore");
                    await showMessage("无法打开", String(error));
                  });
                },
              },
              {
                label: "选择其它应用打开…",
                icon: "external-link" as IconName,
                // 弹出 Windows「打开方式」对话框，临时挑一个程序（例如专门用 WPS 打开这个表）
                onClick: () => void openActiveDocWithExternalApp("choose"),
              },
              {
                label: "在文件资源管理器中显示",
                icon: "folder-open" as IconName,
                onClick: () => void revealItemInDir(doc.filePath!),
              },
            ],
          ]
        : []),
      // 4. 分栏与新窗口
      [
        ...(pane === 0
          ? [
              {
                label: "移到右栏",
                icon: "columns" as IconName,
                onClick: () => useAppStore.getState().moveDoc(doc.id, 1),
              },
            ]
          : [
              {
                label: "移到左栏",
                icon: "columns" as IconName,
                onClick: () => useAppStore.getState().moveDoc(doc.id, 0),
              },
            ]),
        ...(doc.filePath
          ? [
              {
                label: "在新窗口中打开",
                icon: "external-link" as IconName,
                onClick: () => void invoke("open_in_new_window", { path: doc.filePath }),
              },
            ]
          : []),
      ],
      // 5. 保存
      ...(doc.isDirty || !doc.filePath
        ? [
            [
              {
                label: "保存",
                icon: "save" as IconName,
                hint: "Ctrl+S",
                onClick: () => void saveDoc(doc.id),
              },
              {
                label: "另存为…",
                icon: "save-as" as IconName,
                hint: "Ctrl+Shift+S",
                onClick: () => void saveDocAs(doc.id),
              },
            ],
          ]
        : [
            [
              {
                label: "另存为…",
                icon: "save-as" as IconName,
                hint: "Ctrl+Shift+S",
                onClick: () => void saveDocAs(doc.id),
              },
            ],
          ]),
    ];
  };

  /** 处理 Tab 上的 PointerDown，启用平滑拖拽与单击激活 */
  const handleTabPointerDown = (
    e: React.PointerEvent<HTMLDivElement>,
    doc: DocState,
    _index: number,
  ) => {
    // 仅响应鼠标左键，避免拦截右键/中键
    if (e.button !== 0) return;
    // 如果点击在关闭按钮上，不触发拖拽
    if ((e.target as HTMLElement).closest("button")) return;

    const startX = e.clientX;
    const startY = e.clientY;
    const title = getDocTitle(doc, docs);

    const onPointerMove = (moveEvt: PointerEvent) => {
      const dist = Math.hypot(moveEvt.clientX - startX, moveEvt.clientY - startY);
      if (!useTabDragStore.getState().isDragging) {
        if (dist > 4) {
          useTabDragStore.getState().startDrag(doc.id, pane, title, moveEvt.clientX, moveEvt.clientY);
        } else {
          return;
        }
      }

      // 正在拖拽：探测当前指针所在的分栏和插入点
      const elem = document.elementFromPoint(moveEvt.clientX, moveEvt.clientY);
      const barElem = elem?.closest("[data-tabbar-pane]");
      const viewElem = elem?.closest("[data-pane-viewport]");

      let targetPane: 0 | 1 | null = null;
      let targetIndex: number | null = null;

      if (barElem) {
        const p = Number(barElem.getAttribute("data-tabbar-pane")) as 0 | 1;
        targetPane = p;
        const tabElem = elem?.closest("[data-tab-index]");
        if (tabElem) {
          const tIdx = Number(tabElem.getAttribute("data-tab-index"));
          const rect = tabElem.getBoundingClientRect();
          const midX = rect.left + rect.width / 2;
          targetIndex = moveEvt.clientX < midX ? tIdx : tIdx + 1;
        } else {
          const pDocs = useAppStore.getState().docs.filter((d) => (d.pane ?? 0) === p);
          targetIndex = pDocs.length;
        }
      } else if (viewElem) {
        const p = Number(viewElem.getAttribute("data-pane-viewport")) as 0 | 1;
        const pDocs = useAppStore.getState().docs.filter((d) => (d.pane ?? 0) === p);
        targetIndex = pDocs.length;
        if (p === 0 && !useAppStore.getState().layout.split) {
          // 单栏模式：拖到视口右半边即视为放入新开的右栏（与「移到右栏 (双栏并排)」提示一致）。
          // 窗口开启了 dragDropEnabled，Windows 上 HTML5 drag/drop 事件不会派发，
          // 因此这里用指针坐标判断左右半边。
          const rect = viewElem.getBoundingClientRect();
          targetPane = moveEvt.clientX > rect.left + rect.width / 2 ? 1 : 0;
        } else {
          targetPane = p;
        }
      }

      useTabDragStore.getState().updateHover(moveEvt.clientX, moveEvt.clientY, targetPane, targetIndex);
    };

    const onPointerUp = () => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);

      const state = useTabDragStore.getState();
      if (state.isDragging) {
        if (state.docId && state.targetPane !== null) {
          useAppStore.getState().moveDoc(
            state.docId,
            state.targetPane,
            state.targetIndex ?? undefined,
          );
        }
        state.endDrag();
      } else {
        // 未超过阈值，视为正常点击激活
        useAppStore.getState().activateDoc(doc.id, pane);
      }
    };

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  };

  const isDraggingOverThisBar =
    dragStore.isDragging && dragStore.targetPane === pane;

  return (
    <>
      <div
        ref={containerRef}
        role="tablist"
        data-tabbar-pane={pane}
        onWheel={handleWheel}
        className={`flex h-8 shrink-0 items-stretch gap-0.5 overflow-x-auto border-b border-line bg-app px-1.5 pt-0.5 select-none transition-colors ${
          isDraggingOverThisBar ? "bg-accent/5 ring-1 ring-inset ring-accent/20" : ""
        }`}
      >
        {paneDocs.map((doc, index) => {
          const active = doc.id === paneActiveId && !gameOpen;
          const displayTitle = getDocTitle(doc, docs);
          const showLeftIndicator =
            isDraggingOverThisBar && dragStore.targetIndex === index;
          const showRightIndicator =
            isDraggingOverThisBar &&
            dragStore.targetIndex === paneDocs.length &&
            index === paneDocs.length - 1;

          return (
            <div key={doc.id} className="relative flex items-stretch">
              {/* 插入到该标签左侧的指示线 */}
              {showLeftIndicator ? (
                <div className="z-30 my-0.5 w-[2.5px] rounded bg-accent shadow-sm shadow-accent transition-all" />
              ) : null}

              <div
                ref={active ? activeTabRef : null}
                role="tab"
                data-tab-index={index}
                aria-selected={active}
                title={doc.filePath ?? "未保存文档"}
                onPointerDown={(e) => handleTabPointerDown(e, doc, index)}
                onAuxClick={(event) => {
                  if (event.button === 1) void closeDocWithConfirm(doc.id);
                }}
                onContextMenu={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  setContextMenu({ x: event.clientX, y: event.clientY, doc });
                }}
                className={`group relative flex max-w-[170px] shrink-0 cursor-pointer items-center gap-1 self-end rounded-t-md border border-b-0 px-2 py-1 text-[11.5px] transition-colors ${
                  active
                    ? "border-line bg-panel font-medium text-fg"
                    : "border-line-strong/60 bg-hover text-fg/75 hover:bg-active hover:text-fg"
                } ${
                  dragStore.isDragging && dragStore.docId === doc.id
                    ? "opacity-40"
                    : ""
                }`}
                style={active ? { boxShadow: "inset 0 2px 0 0 var(--accent)" } : undefined}
              >
                <Icon
                  name={isPdfDoc(doc) ? "file-pdf" : doc.filePath ? "file-text" : "file-plus"}
                  size={12}
                  className={active ? "text-accent" : "text-muted"}
                />
                <span className="truncate">{displayTitle}</span>
                {doc.isDirty ? (
                  <span className="text-[10px] text-accent" title="有未保存的更改">
                    ●
                  </span>
                ) : null}
                <button
                  type="button"
                  title="关闭标签 (Ctrl+W)"
                  onClick={(event) => {
                    event.stopPropagation();
                    void closeDocWithConfirm(doc.id);
                  }}
                  className={`rounded p-0.5 transition-colors hover:bg-active hover:text-fg ${
                    active ? "text-muted" : "text-muted/70"
                  }`}
                >
                  <Icon name="x" size={11} />
                </button>
              </div>

              {/* 插入到末尾标签右侧的指示线 */}
              {showRightIndicator ? (
                <div className="z-30 my-0.5 w-[2.5px] rounded bg-accent shadow-sm shadow-accent transition-all" />
              ) : null}
            </div>
          );
        })}

        {/* 当分栏内文档为空但被悬停时，显示落点指示 */}
        {isDraggingOverThisBar && paneDocs.length === 0 ? (
          <div className="my-0.5 w-[2.5px] rounded bg-accent shadow-sm shadow-accent" />
        ) : null}

        {/* 江湖独立标签（固定挂在左栏） */}
        {pane === 0 && gameOpen ? (
          <div
            role="tab"
            aria-selected
            title="Markdown 江湖"
            className="group flex max-w-[170px] shrink-0 cursor-pointer items-center gap-1 self-end rounded-t-md border border-line border-b-0 bg-panel px-2 py-1 text-[11.5px] font-medium text-fg"
            style={{ boxShadow: "inset 0 2px 0 0 var(--accent)" }}
          >
            <Icon name="sword" size={12} className="text-accent" />
            <span className="truncate">江湖</span>
            <button
              type="button"
              title="关闭标签"
              onClick={(event) => {
                event.stopPropagation();
                useAppStore.getState().closeGame();
              }}
              className="rounded p-0.5 text-muted transition-colors hover:bg-active hover:text-fg"
            >
              <Icon name="x" size={11} />
            </button>
          </div>
        ) : null}

        {/* 新建标签按钮 */}
        <button
          type="button"
          title="新建或打开文档"
          onClick={(e) => {
            useAppStore.getState().setActivePane(pane);
            const rect = e.currentTarget.getBoundingClientRect();
            setPlusMenu({ x: rect.left, y: rect.bottom + 4 });
          }}
          className="ml-0.5 shrink-0 self-center rounded p-1 text-muted hover:bg-hover hover:text-fg"
        >
          <Icon name="plus" size={13} />
        </button>

        {/* 新建文档下拉菜单 */}
        {plusMenu ? (
          <ContextMenu
            x={plusMenu.x}
            y={plusMenu.y}
            groups={[
              [
                {
                  label: "打开文件…",
                  icon: "folder-open",
                  hint: "Ctrl+O",
                  onClick: () => void openFileDialog(pane),
                },
                {
                  label: "新建 Markdown 文档 (.md)",
                  icon: "file-text",
                  hint: "Ctrl+N",
                  onClick: () => void newDocument("markdown", pane),
                },
                {
                  label: "新建空白文档",
                  icon: "file-plus",
                  onClick: () => void newDocument("blank", pane),
                },
              ],
            ]}
            onClose={() => setPlusMenu(null)}
          />
        ) : null}

        {/* 右侧：双栏分屏切换按钮（在主栏末尾展示） */}
        {pane === 0 ? (
          <button
            type="button"
            title={layout.split ? "关闭双栏 (Ctrl+\\)" : "双栏分屏 (Ctrl+\\)"}
            onClick={() => useAppStore.getState().toggleSplit()}
            className={`ml-auto shrink-0 self-center rounded-md p-1 transition-colors ${
              layout.split ? "bg-hover text-accent" : "text-muted hover:bg-hover hover:text-fg"
            }`}
          >
            <Icon name="columns" size={14} />
          </button>
        ) : null}

        {/* 右键菜单 */}
        {contextMenu ? (
          <ContextMenu
            x={contextMenu.x}
            y={contextMenu.y}
            groups={getContextMenuGroups(contextMenu.doc)}
            onClose={() => setContextMenu(null)}
          />
        ) : null}
      </div>

      {/* 拖拽时的全局跟随小卡片（只在 pane 0 渲染一次即可，避免双栏重复渲染） */}
      {pane === 0 && dragStore.isDragging ? (
        <div
          className="pointer-events-none fixed z-[99999] flex items-center gap-1.5 rounded-md border border-accent/60 bg-panel/95 px-2.5 py-1 text-[12px] font-medium text-fg shadow-2xl backdrop-blur-sm"
          style={{
            left: dragStore.x + 12,
            top: dragStore.y + 12,
            transform: "translate3d(0, 0, 0)",
          }}
        >
          <Icon name="file-text" size={13} className="text-accent" />
          <span className="max-w-[150px] truncate">{dragStore.title}</span>
        </div>
      ) : null}
    </>
  );
}
