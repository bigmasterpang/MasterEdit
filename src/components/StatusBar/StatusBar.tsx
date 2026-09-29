import { useMemo } from "react";
import { useAppStore } from "../../stores/appStore";
import { useSettingsStore } from "../../stores/settingsStore";
import { useUpdateStore } from "../../stores/updateStore";
import { Icon } from "../common/Icon";
import { countWords, formatBytes } from "../../utils/timing";
import { APP_NAME, ENCODINGS, EOL_OPTIONS } from "../../utils/constants";
import { docKindOf, extName, fileName, isDelimitedDoc, isPdfDoc, isSpreadsheetDoc } from "../../utils/filePath";
import { setDocEncoding, setDocEol, toggleDocReadOnly } from "../../utils/fileActions";
import { extractFrontMatter } from "../../utils/frontMatter";

const VIEW_LABEL = {
  preview: "预览",
  source: "源码",
  split: "分屏",
} as const;

const KIND_LABEL = {
  markdown: "Markdown",
  text: "纯文本",
  code: "代码",
  pdf: "PDF",
  delimited: "CSV 表格",
  spreadsheet: "Excel 表格",
} as const;

/** 原始字数：源码中的字符数（不含空白），预览字数：渲染后可见字符数 */
function countChars(text: string): number {
  return text.replace(/\s/g, "").length;
}

export function StatusBar() {
  const doc = useAppStore((s) => s.docs.find((d) => d.id === s.activeId) ?? null);
  const defaultFontSize = useSettingsStore((s) => s.fontSize);
  const viewMode = useAppStore((s) => s.viewMode);
  const updateInfo = useUpdateStore((s) => s.info);
  const docZoomPercent = doc
    ? Math.round(((doc.fontSize ?? defaultFontSize) / defaultFontSize) * 100)
    : 100;

  const stats = useMemo(() => {
    const content = doc?.content ?? "";
    const rawChars = countChars(content);
    const words = countWords(content);
    const lines = content ? content.split("\n").length : 0;
    // 预览字数：去掉 front matter 与 Markdown 语法标记后的可见字符数
    const { body } = extractFrontMatter(content);
    const previewText = body
      .replace(/```[\s\S]*?```/g, (block) => block.replace(/```[^\n]*\n?/g, ""))
      .replace(/`[^`\n]*`/g, "")
      .replace(/!\[[^\]]*\]\([^)]*\)/g, "")
      .replace(/\[([^\]]*)\]\([^)]*\)/g, "$1")
      .replace(/^[ \t]*>[ \t]?/gm, "")
      .replace(/^[ \t]*[-*+][ \t]+/gm, "")
      .replace(/^[ \t]*\d+[.)][ \t]+/gm, "")
      .replace(/^[ \t]*#{1,6}[ \t]+/gm, "")
      .replace(/[*_~^]/g, "")
      .replace(/!\[[^\]]*\]/g, "");
    return {
      rawChars,
      previewChars: countChars(previewText),
      words,
      lines,
    };
  }, [doc?.content]);

  const bytes = useMemo(
    () => (doc ? new TextEncoder().encode(doc.content).length : 0),
    [doc?.content],
  );

  /** 电子表格：只读查看，状态栏改为展示文件信息而非文本统计 */
  const isSheet = isSpreadsheetDoc(doc);
  /** CSV / TSV：同样是表格，字数统计同样没有意义 */
  const isDelimited = isDelimitedDoc(doc);

  return (
    <div className="print-hide flex h-6 shrink-0 items-center gap-3 border-t border-line bg-panel px-3 text-[11px] text-muted">
      <span className="flex min-w-0 items-center gap-1">
        {doc?.isDirty ? <span className="text-accent">●</span> : null}
        <span className="truncate" title={doc?.filePath ?? "未保存"}>
          {doc ? (doc.filePath ? fileName(doc.filePath) : "未命名文档") : APP_NAME}
        </span>
      </span>

      <div className="flex-1" />

      {updateInfo?.hasUpdate ? (
        <button
          type="button"
          onClick={() => useUpdateStore.getState().showDialog()}
          className="flex items-center gap-1 text-accent hover:underline"
          title="发现新版本，点击查看详情"
        >
          <Icon name="download" size={12} />
          发现新版本 {updateInfo.latest}
        </button>
      ) : null}

      {doc ? (
        isPdfDoc(doc) ? (
          <>
            <span title="当前页码">
              第 {doc.pdfCurrentPage ?? 1} / {doc.pdfTotalPages ?? 1} 页
            </span>
            <span title="文件大小">{formatBytes(doc.size)}</span>
            <span title="文件类型">PDF 文档</span>
          </>
        ) : isSheet || isDelimited ? (
          /* 表格类文档（xlsx / CSV）：字数、词数、光标行列对表格没有意义，只显示文件与编辑状态 */
          <>
            <span title="文件大小">{formatBytes(doc.size)}</span>
            <span title="文件类型">
              {KIND_LABEL[docKindOf(doc.filePath)]}
              {extName(doc.filePath ?? "") ? ` · ${extName(doc.filePath ?? "")}` : ""}
            </span>
            {isDelimited ? (
              <span className="text-faint" title="CSV / TSV 走文本管线：双击单元格即可编辑，Ctrl+E 可切到源码">
                表格视图可编辑
              </span>
            ) : null}
            {doc.sheetEdits && doc.sheetEdits.length > 0 ? (
              <span className="text-accent" title="待保存的单元格编辑，按 Ctrl+S 写回文件">
                已修改 {doc.sheetEdits.length} 格
              </span>
            ) : null}
            {doc.sheetStructurePending ? (
              <span
                className="text-accent"
                title="插入/删除行列、工作表增删改复制只在内存中生效，按 Ctrl+S 才写回文件"
              >
                结构已修改（未保存）
              </span>
            ) : null}
            {!doc.readOnly && !isDelimited ? (
              <span className="text-faint" title="双击单元格即可编辑，Ctrl+S 保存">
                可编辑
              </span>
            ) : null}
          </>
        ) : (
          <>
            <span title="光标位置">
              行 {doc.cursorLine}, 列 {doc.cursorCol}
            </span>
            {doc.selectionLength > 0 ? (
              <span title="已选中字符数">选中 {doc.selectionLength}</span>
            ) : null}
            <span title="原始字数：Markdown 源码字符数（不含空白）">
              原始字数 {stats.rawChars}
            </span>
            <span title="预览字数：渲染后可见字符数（不含空白）">
              预览字数 {stats.previewChars}
            </span>
            <span title="词数（中文按字、西文按词）">词数 {stats.words}</span>
            <span className="hidden sm:inline" title="总行数">
              {stats.lines} 行
            </span>
            <span title="文件大小">{doc.size > 0 ? formatBytes(doc.size) : formatBytes(bytes)}</span>
          </>
        )
      ) : (
        <span>就绪</span>
      )}

      {doc && !isPdfDoc(doc) && !isSheet ? (
        <>
          <button
            type="button"
            onClick={() => useAppStore.getState().patchDoc(doc.id, { fontSize: undefined })}
            title="当前文档独立缩放比例（Ctrl+滚轮仅缩放所在窗口文档，点击恢复 100%）"
            className={`rounded px-1 py-0.5 transition-colors hover:bg-hover hover:text-fg ${
              doc.fontSize != null && doc.fontSize !== defaultFontSize ? "font-medium text-accent" : ""
            }`}
          >
            {docZoomPercent}%
          </button>
          <span title="文件类型">
            {KIND_LABEL[docKindOf(doc.filePath)]}
            {extName(doc.filePath ?? "") ? ` · ${extName(doc.filePath ?? "")}` : ""}
          </span>
          <select
            value={doc.encoding}
            onChange={(event) => void setDocEncoding(doc.id, event.target.value)}
            title="文件编码：切换后会按新编码重新读取，保存时按此编码写回"
            className="h-5 rounded border border-line bg-transparent px-1 text-[11px] text-muted hover:text-fg"
          >
            {ENCODINGS.map((enc) => (
              <option key={enc.id} value={enc.id}>
                {enc.label}
              </option>
            ))}
          </select>
          <select
            value={doc.eol}
            onChange={(event) => setDocEol(doc.id, event.target.value)}
            title="换行符：保存时统一转换"
            className="h-5 rounded border border-line bg-transparent px-1 text-[11px] text-muted hover:text-fg"
          >
            {EOL_OPTIONS.map((option) => (
              <option key={option.id} value={option.id}>
                {option.label}
              </option>
            ))}
          </select>
        </>
      ) : !doc ? (
        <span title="编码">UTF-8</span>
      ) : null}
      {doc?.encrypted ? (
        <span
          className="flex items-center gap-1 text-accent"
          title="企业加密文档：已自动解密，保存时按原加密格式写回"
        >
          <Icon name="key" size={11} />
          已解密
        </span>
      ) : null}
      {doc?.readOnly ? (
        <button
          type="button"
          onClick={() => void toggleDocReadOnly(doc.id)}
          className="text-warning hover:underline"
          title="只读文档（超过自动只读阈值或格式不支持写回）：点击切换为可编辑"
        >
          只读 · 点击可编辑
        </button>
      ) : null}
      <span title="当前视图模式">
        {isPdfDoc(doc) ? "PDF 查看器" : isSheet ? "表格查看器" : VIEW_LABEL[viewMode]}
      </span>
    </div>
  );
}
