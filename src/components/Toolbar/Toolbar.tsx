import { Icon, type IconName } from "../common/Icon";
import { DropdownMenu, type MenuGroup } from "../common/DropdownMenu";
import { useAppStore } from "../../stores/appStore";
import { useDialogStore, showMessage, askForm } from "../../stores/dialogStore";
import { useSettingsStore } from "../../stores/settingsStore";
import { useSearchStore } from "../../stores/searchStore";
import type { ThemeMode, ViewMode } from "../../types";
import { getEditorView } from "../../utils/editorBridge";
import {
  insertCallout,
  insertCodeBlock,
  insertDateTime,
  insertHorizontalRule,
  insertMarkdownImage,
  insertMarkdownLink,
  insertMath,
  insertMermaid,
  insertTable,
  insertToc,
  setHeading,
  shiftHeading,
  toggleBold,
  toggleHighlight,
  toggleInlineCode,
  toggleItalic,
  toggleList,
  toggleQuote,
  toggleStrikethrough,
  toggleSubscript,
  toggleSuperscript,
  toggleUnderline,
  transformCase,
} from "../../utils/editorCommands";
import { exportHtmlFile } from "../../utils/exportHtml";
import { exportPngFile } from "../../utils/exportPng";
import { exportDocxFile } from "../../utils/exportDocx";
import { exportPdfFile } from "../../utils/exportPdf";
import {
  newDocument,
  openActiveDocWithExternalApp,
  openFileDialog,
  saveActive,
  saveActiveAs,
} from "../../utils/fileActions";
import { fileName, isDelimitedDoc, isMarkdownDoc, isPdfDoc, isSpreadsheetDoc } from "../../utils/filePath";
import { COMMON_FORMULA_TEMPLATES, sendSheetCommand } from "../../utils/sheetCommands";
import { parseDoc } from "../../utils/markdown";

interface ToolbarProps {
  previewRef: React.RefObject<HTMLDivElement | null>;
  isDark: boolean;
}

function ToolButton({
  icon,
  label,
  onClick,
  active,
  disabled,
}: {
  icon?: IconName;
  label: string;
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
}) {
  return (
    <button
      type="button"
      title={label}
      aria-label={label}
      disabled={disabled}
      onClick={onClick}
      className={`flex h-7 items-center gap-1.5 rounded-md px-2 text-[12px] transition-colors disabled:cursor-not-allowed disabled:opacity-40 ${
        active ? "bg-accent-soft-strong font-medium text-accent" : "text-muted hover:bg-hover hover:text-fg"
      }`}
    >
      {icon ? <Icon name={icon} size={15} /> : null}
    </button>
  );
}

const VIEW_ITEMS: Array<{ mode: ViewMode; icon: IconName; label: string }> = [
  { mode: "preview", icon: "eye", label: "预览模式 (Ctrl+E)" },
  { mode: "source", icon: "code", label: "源码模式 (Ctrl+E)" },
  { mode: "split", icon: "columns", label: "分屏模式 (Ctrl+E)" },
];

/** CSV / TSV：同样的三态，但「预览」实际是表格视图 */
const DELIMITED_VIEW_ITEMS: Array<{ mode: ViewMode; icon: IconName; label: string }> = [
  { mode: "preview", icon: "grid", label: "表格视图 (Ctrl+E)" },
  { mode: "source", icon: "code", label: "源码模式 (Ctrl+E)" },
  { mode: "split", icon: "columns", label: "表格 + 源码 (Ctrl+E)" },
];

const THEME_ORDER: ThemeMode[] = ["light", "dark", "system"];
const THEME_META: Record<ThemeMode, { icon: IconName; label: string }> = {
  light: { icon: "sun", label: "浅色主题" },
  dark: { icon: "moon", label: "深色主题" },
  system: { icon: "monitor", label: "跟随系统" },
};

export function Toolbar({ previewRef, isDark }: ToolbarProps) {
  const doc = useAppStore((s) => s.docs.find((d) => d.id === s.activeId) ?? null);
  const viewMode = useAppStore((s) => s.viewMode);
  const outlineVisible = useAppStore((s) => s.outlineVisible);
  const syncScroll = useAppStore((s) => s.syncScroll);
  const searchVisible = useSearchStore((s) => s.visible);
  const replaceVisible = useSearchStore((s) => s.replaceVisible);
  const theme = useSettingsStore((s) => s.theme);
  const setSetting = useSettingsStore((s) => s.set);

  const hasDoc = Boolean(doc);
  const editable = viewMode !== "preview" && Boolean(doc) && !doc?.readOnly;
  /** 非 Markdown 文档（新建空白文档/代码/纯文本/PDF）屏蔽所有 Markdown 专属操作 */
  const isMarkdown = isMarkdownDoc(doc);
  const isPdf = isPdfDoc(doc);
  /** 电子表格（xlsx/xls/ods）：只读查看，屏蔽查找替换与视图切换 */
  const isSheet = isSpreadsheetDoc(doc);
  /** CSV / TSV：可编辑文本 + 表格视图 */
  const isDelimited = isDelimitedDoc(doc);

  const cycleTheme = () => {
    const idx = THEME_ORDER.indexOf(theme);
    setSetting("theme", THEME_ORDER[(idx + 1) % THEME_ORDER.length]);
  };

  const baseName = doc?.filePath
    ? fileName(doc.filePath).replace(/\.(md|markdown|mdown|mkd|mkdn|txt)$/i, "")
    : "未命名";

  /** 导出前确保处于预览模式并等待渲染完成 */
  const ensurePreview = async (): Promise<HTMLElement | null> => {
    if (!hasDoc) {
      await showMessage("无法导出", "请先打开或新建一个文档。");
      return null;
    }
    if (viewMode === "source") {
      useAppStore.getState().setViewMode("preview");
      await new Promise((resolve) => setTimeout(resolve, 260));
    }
    // 大文档需要手动刷新预览
    if (!previewRef.current) {
      await new Promise((resolve) => setTimeout(resolve, 200));
    }
    if (!previewRef.current) {
      await showMessage("无法导出", "预览尚未就绪，请稍后重试。");
      return null;
    }
    return previewRef.current;
  };

  const handleExport = async (kind: "html" | "png" | "docx" | "pdf") => {
    const node = await ensurePreview();
    if (!node) return;
    if (kind === "html") {
      await exportHtmlFile({ title: baseName, source: node, inlineImages: true });
      return;
    }
    if (kind === "png") {
      await exportPngFile({ source: node, title: baseName, isDark });
      return;
    }
    if (kind === "docx") {
      await exportDocxFile({
        source: doc?.content ?? "",
        tokens: parseDoc(doc?.content ?? ""),
        docPath: doc?.filePath ?? null,
        title: baseName,
      });
      return;
    }
    await exportPdfFile({
      title: baseName,
      beforePrint: async () => {
        useSearchStore.getState().close();
      },
    });
  };

  const headingMenu: MenuGroup[] = [
    {
      title: "标题级别",
      items: [1, 2, 3, 4, 5, 6].map((level) => ({
        label: `H${level} 标题`,
        hint: `${"#".repeat(level)} Ctrl+${level}`,
        onClick: () => setHeading(level),
      })),
    },
    {
      items: [
        { label: "普通段落", hint: "Ctrl+0", onClick: () => setHeading(0) },
        { label: "提升一级", hint: "Ctrl+Alt+↑", onClick: () => shiftHeading(-1) },
        { label: "降低一级", hint: "Ctrl+Alt+↓", onClick: () => shiftHeading(1) },
      ],
    },
  ];

  const listMenu: MenuGroup[] = [
    {
      title: "列表",
      items: [
        { icon: "list", label: "无序列表", hint: "Ctrl+Shift+8", onClick: () => toggleList("bullet") },
        { icon: "list", label: "有序列表", hint: "Ctrl+Shift+7", onClick: () => toggleList("ordered") },
        { icon: "check", label: "任务列表", hint: "Ctrl+Shift+9", onClick: () => toggleList("task") },
        { icon: "chevron-right", label: "引用块", hint: "Ctrl+Shift+Q", onClick: () => toggleQuote() },
      ],
    },
    {
      title: "提示块",
      items: [
        { label: "提示 (NOTE)", onClick: () => insertCallout("note") },
        { label: "技巧 (TIP)", onClick: () => insertCallout("tip") },
        { label: "重要 (IMPORTANT)", onClick: () => insertCallout("important") },
        { label: "警告 (WARNING)", onClick: () => insertCallout("warning") },
        { label: "注意 (CAUTION)", onClick: () => insertCallout("caution") },
      ],
    },
  ];

  const insertMenu: MenuGroup[] = [
    {
      title: "插入",
      items: [
        {
          icon: "image",
          label: "链接",
          hint: "Ctrl+K",
          onClick: () => void promptInsert("link"),
        },
        {
          icon: "image",
          label: "图片",
          hint: "Ctrl+Shift+I",
          onClick: () => void promptInsert("image"),
        },
        { icon: "columns", label: "表格 3 × 3", hint: "Ctrl+Shift+T", onClick: () => insertTable(3, 3) },
        { icon: "code", label: "代码块", hint: "Ctrl+Shift+C", onClick: () => insertCodeBlock() },
        { icon: "minus", label: "分割线", hint: "Ctrl+Shift+H", onClick: () => insertHorizontalRule() },
        { icon: "clock", label: "日期时间", onClick: () => insertDateTime() },
        { icon: "list", label: "目录 (TOC)", hint: "Ctrl+Shift+O", onClick: () => insertToc() },
      ],
    },
    {
      title: "数学公式",
      items: [
        { label: "行内公式  $E = mc^2$", onClick: () => insertMath(false) },
        { label: "块级公式  $$ … $$", onClick: () => insertMath(true) },
      ],
    },
    {
      title: "图表",
      items: [
        { label: "Mermaid 流程图", icon: "columns", onClick: () => insertMermaid() },
      ],
    },
    {
      title: "更多表格",
      items: [
        { label: "2 列 × 3 行", onClick: () => insertTable(2, 3) },
        { label: "4 列 × 4 行", onClick: () => insertTable(4, 4) },
        { label: "5 列 × 3 行", onClick: () => insertTable(5, 3) },
      ],
    },
  ];

  const formatMenu: MenuGroup[] = [
    {
      title: "文本格式",
      items: [
        { icon: "bold", label: "粗体", hint: "Ctrl+B", onClick: toggleBold },
        { icon: "italic", label: "斜体", hint: "Ctrl+I", onClick: toggleItalic },
        { icon: "minus", label: "删除线", hint: "Alt+Shift+5", onClick: toggleStrikethrough },
        { label: "高亮", hint: "Ctrl+Shift+M", onClick: toggleHighlight },
        { label: "下划线", hint: "Alt+Shift+U", onClick: toggleUnderline },
        { icon: "code", label: "行内代码", hint: "Ctrl+`", onClick: toggleInlineCode },
        { icon: "link", label: "链接", hint: "Ctrl+K", onClick: () => void promptInsert("link") },
        { label: "上标 x²", onClick: toggleSuperscript },
        { label: "下标 x₂", onClick: toggleSubscript },
      ],
    },
    {
      title: "大小写",
      items: [
        { label: "转换为大写", onClick: () => transformCase("upper") },
        { label: "转换为小写", onClick: () => transformCase("lower") },
        { label: "首字母大写", onClick: () => transformCase("title") },
      ],
    },
  ];

  const exportMenu: MenuGroup[] = [
    {
      title: "导出",
      items: [
        { icon: "download", label: "HTML 单文件", hint: ".html", onClick: () => void handleExport("html") },
        { icon: "image", label: "图片 (PNG)", hint: ".png", onClick: () => void handleExport("png") },
        { icon: "file-text", label: "Word 文档", hint: ".docx", onClick: () => void handleExport("docx") },
        { icon: "file-text", label: "PDF 文档", hint: ".pdf", onClick: () => void handleExport("pdf") },
      ],
    },
  ];

  return (
    <div className="print-hide flex h-11 shrink-0 items-center gap-1 border-b border-line bg-panel px-2">
      <ToolButton icon="file-plus" label="新建 (Ctrl+N)" onClick={() => void newDocument()} />
      <ToolButton icon="folder-open" label="打开 (Ctrl+O)" onClick={() => void openFileDialog()} />
      <ToolButton icon="save" label="保存 (Ctrl+S)" disabled={!hasDoc} onClick={() => void saveActive()} />
      <ToolButton
        icon="save-as"
        label="另存为 (Ctrl+Shift+S)"
        disabled={!hasDoc}
        onClick={() => void saveActiveAs()}
      />
      {/* 用本地其它应用打开当前文档（Excel / CSV 需要完整编辑能力时最顺手；加密文档由系统驱动透明解密） */}
      <DropdownMenu
        icon="external-link"
        label="打开方式"
        title="用本地其它应用打开当前文档（打开的是磁盘上的版本）"
        disabled={!doc?.filePath}
        groups={[
          {
            items: [
              {
                label: "用系统默认程序打开",
                onClick: () => void openActiveDocWithExternalApp("default"),
              },
              {
                label: "选择其它应用…",
                onClick: () => void openActiveDocWithExternalApp("choose"),
              },
            ],
          },
        ]}
      />

      <Divider />

      {isMarkdown ? (
        <>
          <div className="flex items-center gap-0.5 rounded-md bg-app p-0.5">
            {VIEW_ITEMS.map((item) => (
              <ToolButton
                key={item.mode}
                icon={item.icon}
                label={item.label}
                active={viewMode === item.mode}
                disabled={!hasDoc}
                onClick={() => useAppStore.getState().setViewMode(item.mode)}
              />
            ))}
          </div>

          <Divider />

          <DropdownMenu
            icon="file-text"
            label="标题"
            title="标题级别"
            groups={headingMenu}
            disabled={!editable}
          />
          <DropdownMenu
            icon="list"
            label="列表"
            title="列表与提示块"
            groups={listMenu}
            disabled={!editable}
            width={230}
          />
          <DropdownMenu
            icon="plus"
            label="插入"
            title="插入元素"
            groups={insertMenu}
            disabled={!editable}
            width={230}
          />
          <DropdownMenu
            icon="bold"
            label="格式"
            title="文本格式"
            groups={formatMenu}
            disabled={!editable}
            width={230}
          />

          <Divider />
        </>
      ) : null}

      {/* CSV / TSV：表格 / 源码 / 分屏 三态切换（Markdown 分支之外单独提供） */}
      {isDelimited ? (
        <>
          <div className="flex items-center gap-0.5 rounded-md bg-app p-0.5">
            {DELIMITED_VIEW_ITEMS.map((item) => (
              <ToolButton
                key={item.mode}
                icon={item.icon}
                label={item.label}
                active={viewMode === item.mode}
                disabled={!hasDoc}
                onClick={() => useAppStore.getState().setViewMode(item.mode)}
              />
            ))}
          </div>
          <Divider />
        </>
      ) : null}

      {/* 表格类文档：自动调整（列/行/两者）、插入常用公式、冻结窗格 */}
      {isSheet || isDelimited ? (
        <>
          <DropdownMenu
            icon="columns"
            label="自动调整"
            title="按内容自动调整表格的列宽 / 行高"
            disabled={!hasDoc}
            groups={[
              {
                items: [
                  {
                    label: "自动调整列宽",
                    onClick: () => sendSheetCommand({ kind: "autoFit", target: "columns" }),
                  },
                  {
                    label: "自动调整行高",
                    onClick: () => sendSheetCommand({ kind: "autoFit", target: "rows" }),
                  },
                  {
                    label: "列宽与行高都调整",
                    onClick: () => sendSheetCommand({ kind: "autoFit", target: "both" }),
                  },
                  {
                    label: "自动换行并调整行高",
                    onClick: () => sendSheetCommand({ kind: "autoFit", target: "wrap" }),
                  },
                ],
              },
            ]}
          />
          {isSheet ? (
          <DropdownMenu
            icon="fx"
            label="插入公式"
            title="把常用公式插入当前单元格（当前选区会作为公式参数）"
            disabled={!hasDoc || doc?.readOnly}
            groups={[
              {
                items: COMMON_FORMULA_TEMPLATES.map((item) => ({
                  label: item.label,
                  onClick: () => sendSheetCommand({ kind: "insertFormula", template: item.template }),
                })),
              },
            ]}
          />
          ) : null}
          {isSheet ? (
            <DropdownMenu
              icon="freeze"
              label="冻结"
              title="冻结首行 / 首列，滚动时始终可见"
              disabled={!hasDoc}
              groups={[
                {
                  items: [
                    { label: "不冻结", onClick: () => sendSheetCommand({ kind: "freeze", rows: 0, cols: 0 }) },
                    { label: "冻结首行", onClick: () => sendSheetCommand({ kind: "freeze", rows: 1, cols: 0 }) },
                    { label: "冻结首列", onClick: () => sendSheetCommand({ kind: "freeze", rows: 0, cols: 1 }) },
                    {
                      label: "冻结首行和首列",
                      onClick: () => sendSheetCommand({ kind: "freeze", rows: 1, cols: 1 }),
                    },
                  ],
                },
              ]}
            />
          ) : null}
          <Divider />
        </>
      ) : null}

      <ToolButton
        icon="search"
        label="查找 (Ctrl+F)"
        disabled={!hasDoc || isSheet}
        active={searchVisible && !replaceVisible}
        onClick={() => useSearchStore.getState().open(false)}
      />
      <ToolButton
        icon="replace"
        label="替换 (Ctrl+H)"
        disabled={!hasDoc || isPdf || isSheet}
        active={searchVisible && replaceVisible}
        onClick={() => useSearchStore.getState().open(true)}
      />
      <ToolButton
        icon="sidebar"
        label="侧栏 (文件与大纲)"
        active={outlineVisible}
        onClick={() => useAppStore.getState().toggleOutline()}
      />
      {viewMode === "split" && isMarkdown ? (
        <ToolButton
          icon="arrow-up-down"
          label={syncScroll ? "滚动同步：开" : "滚动同步：关"}
          active={syncScroll}
          onClick={() => useAppStore.getState().toggleSyncScroll()}
        />
      ) : null}

      <div className="flex-1" />

      {isMarkdown ? (
        <DropdownMenu
          icon="download"
          label="导出"
          title="导出文档"
          groups={exportMenu}
          disabled={!hasDoc}
          align="right"
          width={200}
        />
      ) : null}
      <ToolButton
        icon={THEME_META[theme].icon}
        label={`${THEME_META[theme].label}（点击切换）`}
        onClick={cycleTheme}
      />
      <ToolButton
        icon="sword"
        label="江湖"
        onClick={() => {
          useAppStore.getState().openGame();
          void import("../../wuxia/store").then((m) => m.useWuxiaStore.getState().ensureLoaded());
        }}
      />
      <ToolButton
        icon="keyboard"
        label="快捷键 (F1)"
        onClick={() => useDialogStore.getState().setShortcutsVisible(true)}
      />
      <ToolButton
        icon="settings"
        label="设置 (Ctrl+,)"
        onClick={() => useDialogStore.getState().setSettingsVisible(true)}
      />
    </div>
  );
}

function Divider() {
  return <div className="mx-1 h-5 w-px bg-line" />;
}

/** 链接 / 图片插入：先弹出表单填写文字与地址，再写入编辑器 */
export async function promptInsert(kind: "link" | "image"): Promise<void> {
  const selected = getSelectedText();
  const values = await askForm({
    title: kind === "link" ? "插入链接" : "插入图片",
    confirmText: "插入",
    fields:
      kind === "link"
        ? [
            { key: "text", label: "显示文字", value: selected, placeholder: "链接文字" },
            {
              key: "url",
              label: "链接地址",
              value: selected && /^https?:|^mailto:/i.test(selected) ? selected : "https://",
              placeholder: "https://example.com",
            },
          ]
        : [
            { key: "alt", label: "图片说明", value: selected, placeholder: "图片描述" },
            {
              key: "src",
              label: "图片路径或链接",
              placeholder: "assets/image.png 或 https://…",
            },
          ],
  });
  if (!values) return;
  if (kind === "link") {
    if (!values.url?.trim()) {
      await showMessage("无法插入链接", "请填写链接地址。");
      return;
    }
    insertMarkdownLink(values.text ?? "", values.url);
  } else {
    if (!values.src?.trim()) {
      await showMessage("无法插入图片", "请填写图片路径或链接。");
      return;
    }
    insertMarkdownImage(values.alt ?? "", values.src);
  }
}

/** 取编辑器或页面选区中的纯文本，用于表单预填 */
function getSelectedText(): string {
  const view = getEditorView();
  if (view) {
    const sel = view.state.selection.main;
    if (!sel.empty) {
      const text = view.state.sliceDoc(sel.from, sel.to);
      if (text.length <= 200 && !text.includes("\n")) return text;
    }
  }
  const page = window.getSelection()?.toString() ?? "";
  return page.length <= 200 && !page.includes("\n") ? page : "";
}
