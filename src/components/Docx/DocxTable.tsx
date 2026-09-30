/**
 * 表格渲染：真实 `<table>` + `colgroup`，支持横向合并（`gridSpan` → `colSpan`）、
 * 纵向合并（`vMerge` → 自己配对算 `rowspan`）、边框、底纹、垂直对齐、嵌套表格。
 *
 * 三条必须守住的约定：
 *  1. **纵向合并的 `continue` 单元格不输出 `<td>`** —— 否则表格会多出一整行格子；
 *  2. **表格整体对齐按 `w:tblPr/w:jc` 还原**：后端把它单独解析成 `table.align`，
 *     与「单元格内段落」的 `jc` 是两条独立的路（样本里 200 多处 `jc=center` 属于后者，
 *     不会影响表格本身）；有显式宽度 + fixed 布局时 `margin auto` 才生效。
 *  3. 单元格里用**同一套块渲染**（多段落 / 嵌套表格），深度上限 6 层。
 */
import { memo, useMemo, type CSSProperties, type ReactNode } from "react";
import type { DocBlock, DocTable, DocTableCell } from "../../types";
import { BLOCK_MARGIN_Y, CELL_PADDING_X, CELL_PADDING_Y, hexColor, round2 } from "./docxStyle";
import { buildTableGrid, type TableGridEntry } from "./docxGrid";
import { MAX_BLOCK_DEPTH, type DocxRenderContext } from "./docxRender";

/** 单元格边框颜色：块模型只给了「有没有边框」，颜色用中性灰（深色主题下也看得清） */
const BORDER_COLOR = "#9aa0aa";
/** 表头行底色（`w:tblHeader`） */
const HEADER_BACKGROUND = "rgba(127, 140, 160, 0.16)";

export interface DocxTableProps {
  table: DocTable;
  ctx: DocxRenderContext;
  /** 嵌套深度（顶层表格 = 0） */
  depth: number;
  /** 表格所在版心的可用宽（px）：算不出单元格宽度时兜底 */
  contentWidth: number;
  /**
   * 单元格内容的块渲染器（由 DocxBlock 注入，避免两个模块循环依赖）。
   * 第三个参数是**单元格内宽**：单元格里的图片要按它算显示盒子，与高度估算保持同一个口径。
   */
  renderBlocks: (blocks: DocBlock[], depth: number, cellWidth: number) => ReactNode;
}

/** 单元格四边边框的内联样式（`border-collapse: collapse` 下相邻边只画一次） */
function cellBorderStyle(cell: DocTableCell): CSSProperties {
  const line = `1px solid ${BORDER_COLOR}`;
  return {
    borderTop: cell.borders.top ? line : undefined,
    borderLeft: cell.borders.left ? line : undefined,
    borderBottom: cell.borders.bottom ? line : undefined,
    borderRight: cell.borders.right ? line : undefined,
  };
}

/** 单元格基础样式：内边距、垂直对齐、底纹 */
function cellBaseStyle(cell: DocTableCell): CSSProperties {
  const style: CSSProperties = {
    padding: `${CELL_PADDING_Y}px ${CELL_PADDING_X}px`,
    verticalAlign: cell.vAlign ?? "top",
    // 单元格默认左对齐；段落自己的 text-align 会覆盖它（样本里的居中都在这一层）
    textAlign: "left",
    wordBreak: "break-word",
    overflowWrap: "anywhere",
  };
  const shading = hexColor(cell.shading);
  if (shading) style.backgroundColor = shading;
  return style;
}

/** 表头行的前几行（连续 `header: true` 的前缀进 `<thead>`） */
function headerRowCount(table: DocTable): number {
  let count = 0;
  for (const row of table.rows) {
    if (!row.header) break;
    count += 1;
  }
  return count;
}

function TableRow({
  entries,
  header,
  heightPx,
  ctx,
  depth,
  columnWidths,
  contentWidth,
  renderBlocks,
}: {
  entries: TableGridEntry[];
  header: boolean;
  heightPx: number | null;
  ctx: DocxRenderContext;
  depth: number;
  columnWidths: number[];
  contentWidth: number;
  renderBlocks: (blocks: DocBlock[], depth: number, cellWidth: number) => ReactNode;
}) {
  const rowStyle: CSSProperties | undefined =
    heightPx !== null && heightPx > 0 ? { height: `${round2(heightPx * ctx.scale)}px` } : undefined;
  return (
    <tr data-docx-row style={rowStyle} data-docx-header-row={header ? "true" : undefined}>
      {entries.map((entry, cellIndex) => {
        const style: CSSProperties = { ...cellBaseStyle(entry.cell), ...cellBorderStyle(entry.cell) };
        if (header) {
          style.backgroundColor = style.backgroundColor ?? HEADER_BACKGROUND;
          style.fontWeight = 600;
        }
        /**
         * 单元格内宽：优先用单元格自己的 `widthPx`，否则按它跨越的列宽求和；
         * 都拿不到就用表格版心宽（与 `estimateCellHeight` 的口径一致）。
         */
        let cellBox = entry.cell.widthPx !== null && entry.cell.widthPx > 0 ? entry.cell.widthPx * ctx.scale : 0;
        if (cellBox <= 0) {
          for (let column = entry.colStart; column < entry.colStart + entry.colSpan; column += 1) {
            cellBox += columnWidths[column] ?? 0;
          }
        }
        if (cellBox <= 0) cellBox = contentWidth;
        const cellInnerWidth = Math.max(40, cellBox - CELL_PADDING_X * 2);
        return (
          <td
            key={cellIndex}
            data-docx-cell
            data-docx-col={entry.colStart}
            data-docx-continuation={entry.continuation ? "true" : undefined}
            data-docx-placeholder={entry.placeholder ? "true" : undefined}
            colSpan={entry.colSpan > 1 ? entry.colSpan : undefined}
            rowSpan={entry.rowSpan > 1 ? entry.rowSpan : undefined}
            style={style}
          >
            {renderBlocks(entry.cell.blocks, depth, cellInnerWidth)}
          </td>
        );
      })}
    </tr>
  );
}

export const DocxTable = memo(function DocxTable({
  table,
  ctx,
  depth,
  contentWidth,
  renderBlocks,
}: DocxTableProps) {
  /** 网格配对（rowspan 计算）只在表格数据变化时重算 */
  const grid = useMemo(() => buildTableGrid(table), [table]);
  const columnCount = Math.max(1, grid.columnCount, table.columns.length);
  const headerCount = headerRowCount(table);
  /**
   * 按行分好、**并剔除纵向延续格**（`continuation`）——那些格子在 OOXML 里存在
   * （后端确认：`continue` 单元格一个都不省），但在 HTML 里必须靠起点格子的 `rowSpan` 表示，
   * 多输出一个 `<td>` 就会让整行往右挤一格。
   */
  const rows = useMemo(() => {
    const grouped: TableGridEntry[][] = table.rows.map(() => []);
    for (const entry of grid.entries) {
      if (entry.continuation) continue;
      grouped[entry.rowIndex]?.push(entry);
    }
    return grouped;
  }, [grid, table.rows]);

  /** 列宽：优先 `columns`（w:tblGrid），缺失时均分 —— `table-layout: fixed` 下必须有宽度 */
  const columnWidths = useMemo(() => {
    const widths: number[] = [];
    for (let c = 0; c < columnCount; c += 1) {
      const value = table.columns[c];
      widths.push(value !== undefined && value > 0 ? round2(value * ctx.scale) : 0);
    }
    const known = widths.reduce((sum, value) => sum + value, 0);
    if (known <= 0) {
      // 一个列宽都没给：均分（用后端给的表格总宽，没有就留空让浏览器自己分）
      const total = table.widthPx !== null && table.widthPx > 0 ? table.widthPx * ctx.scale : 0;
      return widths.map(() => (total > 0 ? round2(total / columnCount) : 0));
    }
    return widths;
  }, [columnCount, ctx.scale, table.columns, table.widthPx]);

  const knownWidth = columnWidths.reduce((sum, value) => sum + value, 0);
  const tableStyle: CSSProperties = {
    borderCollapse: "collapse",
    // 列宽来自 w:tblGrid：固定布局才能保证「列宽 + 合并」按文档来
    tableLayout: knownWidth > 0 ? "fixed" : "auto",
    width: knownWidth > 0 ? `${round2(knownWidth)}px` : undefined,
    maxWidth: "100%",
    // 表格整体对齐：还原 w:tblPr/w:jc（后端单独解析成 table.align，与单元格内段落 jc 不混淆）
    marginLeft: table.align === "center" || table.align === "right" ? "auto" : 0,
    marginRight: table.align === "center" ? "auto" : 0,
  };

  const renderRow = (rowIndex: number) => {
    const row = table.rows[rowIndex];
    return (
      <TableRow
        key={rowIndex}
        entries={rows[rowIndex] ?? []}
        header={row.header}
        heightPx={row.heightPx}
        ctx={ctx}
        depth={depth + 1}
        columnWidths={columnWidths}
        contentWidth={contentWidth}
        renderBlocks={renderBlocks}
      />
    );
  };

  const body: ReactNode[] = [];
  for (let r = headerCount; r < table.rows.length; r += 1) body.push(renderRow(r));
  const head: ReactNode[] = [];
  for (let r = 0; r < headerCount; r += 1) head.push(renderRow(r));

  if (depth >= MAX_BLOCK_DEPTH) {
    return (
      <div
        data-docx-table
        data-docx-depth-limit="true"
        className="my-1 rounded border border-line bg-panel px-2 py-1 text-[11px] text-muted"
      >
        嵌套表格层数过深（超过 {MAX_BLOCK_DEPTH} 层），已停止继续展开。
      </div>
    );
  }

  return (
    <div
      data-docx-table
      style={{
        paddingTop: `${BLOCK_MARGIN_Y.table}px`,
        paddingBottom: `${BLOCK_MARGIN_Y.table}px`,
        // 表格块自身也是左对齐的块级容器
        display: "flex",
        justifyContent: "flex-start",
        overflowX: "auto",
      }}
      onContextMenu={(event) => {
        if (!ctx.onTableContextMenu) return;
        event.preventDefault();
        ctx.onTableContextMenu({ clientX: event.clientX, clientY: event.clientY }, table);
      }}
    >
      <table style={tableStyle} data-docx-columns={columnCount}>
        <colgroup>
          {columnWidths.map((width, index) => (
            <col key={index} style={width > 0 ? { width: `${width}px` } : undefined} />
          ))}
        </colgroup>
        {head.length > 0 ? <thead>{head}</thead> : null}
        <tbody>{body}</tbody>
      </table>
    </div>
  );
});
