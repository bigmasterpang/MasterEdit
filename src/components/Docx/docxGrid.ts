/**
 * 表格网格配对（纯函数，无 React 依赖）：把 OOXML 的 `gridSpan` / `vMerge` 还原成
 * 「每个格子占哪几列、跨几行」，供 `<table>` 渲染与测试断言。
 *
 * ── 为什么需要自己配对 ────────────────────────────────────────────────
 * OOXML 里纵向合并是「起点 + 延续」两段信息：
 *   · 起点：`<w:vMerge w:val="restart"/>`（后端 `vMerge = "restart"`）
 *   · 延续：`<w:vMerge/>`（后端 `vMerge = "continue"`）—— 这些格子在 HTML 里
 *     **不能再输出 `<td>`**，而是给上面那个起点格子的 `rowSpan` 加一；
 *   横向合并 `gridSpan > 1` 直接映射成 `colSpan`。
 *
 * ── 配对算法（单遍，按行推进）──────────────────────────────────────────
 * 维护一张「仍开放的纵向合并」表：列号 → 合并记录（同一条记录的每一列都指向它）。
 *  1. 行内从左往右走，`col` 是下一个可落格子的列号；
 *  2. 遇到 `continue` 且该列正好是某条合并的**起点列** → 认领：起点格子的 `rowSpan += 1`，
 *     本格不产生 `<td>`（`continuation: true`）；
 *  3. 遇到普通格子 / `restart` → 在 `col` 落一个 `<td>`；若 `col` 被上方合并占着，
 *     说明那条合并**本行没有 continue 认领 = 到此结束**，先把它关掉再落格子
 *     （这样即使生产工具省略了 continue 格，网格也不会错位）；
 *  4. `restart` 登记一条新合并记录（本行结束时不会被误关）；
 *  5. 行末把「从上方延续进来、本行却没认领」的合并全部关掉 —— 纵向合并不跨过空白行。
 *
 * `colSpan` 一律取「起点格子」的 `gridSpan`：延续格子的 `gridSpan` 与起点不一致
 * （少见，属畸形文档）时以起点为准，避免网格错位。
 */
import type { DocParagraph, DocBlock, DocTable, DocTableCell, DocTableRow } from "../../types";

/** 一个落到网格上的格子 */
export interface TableGridEntry {
  /** 所属行对象（引用相等；高度估算按行过滤时用它） */
  row: DocTableRow;
  /** 行下标（0 起） */
  rowIndex: number;
  /** 该格子所在的起始列（0 起，已考虑 gridSpan 与纵向合并） */
  colStart: number;
  /** 横向跨越的列数（≥1）= HTML `colSpan` */
  colSpan: number;
  /** 纵向跨越的行数（≥1）= HTML `rowSpan`；只有起点格子有意义 */
  rowSpan: number;
  /** 是否为纵向合并的"延续格"：**不渲染 `<td>`** */
  continuation: boolean;
  /** 是否为「补齐列数」的空占位格（生成器省略了单元格时用；复制文本时忽略） */
  placeholder: boolean;
  cell: DocTableCell;
}

/** 补齐列数用的空单元格（不带边框、不带内容） */
const EMPTY_CELL: DocTableCell = {
  blocks: [],
  text: "",
  gridSpan: 1,
  vMerge: "none",
  widthPx: null,
  shading: null,
  vAlign: null,
  borders: { top: false, left: false, bottom: false, right: false },
};

export interface TableGrid {
  entries: TableGridEntry[];
  /** 网格总列数（`columns` 与实际占位的最大值） */
  columnCount: number;
}

/** 一条进行中的纵向合并记录 */
interface OpenMerge {
  startRow: number;
  colStart: number;
  colSpan: number;
  rowSpan: number;
  /** 起点格子的网格项：续格认领时要同步它的 rowSpan（HTML 只看起点那一格） */
  startEntry: TableGridEntry;
}

export function buildTableGrid(table: DocTable): TableGrid {
  const entries: TableGridEntry[] = [];
  /** 仍"开放"的纵向合并：列号 → 记录 */
  const openByCol = new Map<number, OpenMerge>();
  let columnCount = Math.max(0, table.columns.length);

  /** 结束一条合并记录：清掉它占住的所有列 */
  const closeMerge = (open: OpenMerge): void => {
    for (let c = open.colStart; c < open.colStart + open.colSpan; c += 1) {
      if (openByCol.get(c) === open) openByCol.delete(c);
    }
  };

  table.rows.forEach((row, rowIndex) => {
    /** 本行认领（或新建）的合并：行末据此判断哪些合并已经结束 */
    const touched = new Set<OpenMerge>();
    let col = 0;
    let cellIndex = 0;

    while (cellIndex < row.cells.length) {
      const cell = row.cells[cellIndex];
      const span = Math.max(1, cell.gridSpan);

      if (cell.vMerge === "continue") {
        const open = openByCol.get(col);
        if (open && open.colStart === col) {
          // 认领：起点格子的 rowSpan +1，本格不产生 <td>
          open.rowSpan += 1;
          open.startEntry.rowSpan = open.rowSpan;
          touched.add(open);
          entries.push({
            row,
            rowIndex,
            colStart: col,
            colSpan: open.colSpan,
            rowSpan: open.rowSpan,
            continuation: true,
            placeholder: false,
            cell,
          });
          // 步进按「合并自身的列宽」而不是本格的 gridSpan：畸形文档里两者可能不一致，
          // 按本格步进会让后面的格子撞上刚认领的合并列，把合并误判成结束
          col = Math.max(col + span, open.colStart + open.colSpan);
          cellIndex += 1;
          if (col > columnCount) columnCount = col;
          continue;
        }
        // 找不到起点（畸形文档）：退化成普通格子渲染，**内容绝不静默丢失**
      }

      // 普通格子：先让开"上方合并占住、本行又没认领"的列（那些合并到此结束）
      while (openByCol.has(col)) {
        const stale = openByCol.get(col);
        if (!stale || stale.startRow >= rowIndex) break;
        closeMerge(stale);
      }
      const entry: TableGridEntry = {
        row,
        rowIndex,
        colStart: col,
        colSpan: span,
        rowSpan: 1,
        continuation: false,
        placeholder: false,
        cell,
      };
      entries.push(entry);
      if (cell.vMerge === "restart") {
        const open: OpenMerge = {
          startRow: rowIndex,
          colStart: col,
          colSpan: span,
          rowSpan: 1,
          startEntry: entry,
        };
        for (let c = col; c < col + span; c += 1) openByCol.set(c, open);
        touched.add(open);
      }
      col += span;
      cellIndex += 1;
      if (col > columnCount) columnCount = col;
    }

    // 少数生成器会整格省略 continue 单元格：本行总 gridSpan 少于列数时补空占位格，
    // 否则整行会向左错位（被占住的列交给上面的合并，不补）
    while (col < columnCount && !openByCol.has(col)) {
      entries.push({
        row,
        rowIndex,
        colStart: col,
        colSpan: 1,
        rowSpan: 1,
        continuation: false,
        placeholder: true,
        cell: EMPTY_CELL,
      });
      col += 1;
    }

    // 行末：从上方延续进来、本行没认领的合并 → 结束（不跨空白行延续）
    for (const open of new Set(openByCol.values())) {
      if (open.startRow >= rowIndex) continue;
      if (!touched.has(open)) closeMerge(open);
    }
  });

  return { entries, columnCount };
}

/** 某一行的可见格子（不含纵向延续格，含补齐列数的空占位格） */
export function visibleCells(grid: TableGrid, rowIndex: number): TableGridEntry[] {
  return grid.entries.filter((entry) => entry.rowIndex === rowIndex && !entry.continuation);
}

/**
 * 表格 → 制表符分隔文本（可直接粘进 Excel）。
 * 单元格内容优先用**单元格内的块**渲染成文本（这样单元格里的列表段也带编号），
 * 块为空时才退回后端给的 `cell.text`；
 * 单元格文本里的换行换成空格（否则会破坏「一行 = 一条记录」的语义）；
 * 纵向延续的格子跳过（内容已经在起点格子里了），补齐列数的空占位格也跳过。
 */
export function tableToTsv(table: DocTable): string {
  const grid = buildTableGrid(table);
  return table.rows
    .map((_row, rowIndex) =>
      visibleCells(grid, rowIndex)
        .filter((entry) => !entry.placeholder)
        .map((entry) => cellPlainText(entry.cell).replace(/\r?\n/g, " ").replace(/\t/g, " "))
        .join("\t"),
    )
    .join("\n");
}

/**
 * 段落复制成文本时**带上列表前缀**。
 *
 * Word 里复制列表项是带编号的（`一、xxx`、`1. xxx`），而后端的 `paragraph.text` 只有正文、
 * 前缀单独放在 `list.prefix` 里，所以这里按 `suffix` 拼回去：
 * `tab` → 制表符、`space` → 空格、`nothing` → 直接相接（`numFmt=none` 时 prefix 为空，不加）。
 */
export function paragraphPlainText(block: DocParagraph): string {
  const list = block.list;
  if (!list || !list.prefix) return block.text;
  const separator = list.suffix === "tab" ? "\t" : list.suffix === "space" ? " " : "";
  return `${list.prefix}${separator}${block.text}`;
}

/**
 * 单元格 → 纯文本：优先用单元格内的块（列表段带编号、嵌套表格按 TSV 展开），
 * 块为空时退回 `cell.text`（延续格 / 空格的 text 本来就是空的）。
 */
export function cellPlainText(cell: DocTableCell): string {
  if (cell.blocks.length === 0) return cell.text;
  return cell.blocks
    .map(blockPlainText)
    .filter((text): text is string => text !== null)
    .join("\n");
}

/**
 * 单块 → 文本（复制用）。`null` 表示这个块不产生文本行（分页符），
 * 这样「复制全部文本」与「单元格文本」共用同一套规则，不会各写一份。
 */
function blockPlainText(block: DocBlock): string | null {
  switch (block.kind) {
    case "paragraph":
      return paragraphPlainText(block);
    case "table":
      return tableToTsv(block);
    case "image":
      return `[图片${block.alt ? `：${block.alt}` : block.name ? `：${block.name}` : ""}]`;
    case "unsupported":
      return `[${block.label}]`;
    default:
      // 分页符不产生文本行（否则会多出空行）
      return null;
  }
}

/**
 * 把块列表拍平成纯文本（「复制全部文本」用）。
 * 段落之间用 `\n` 连接；表格按制表符文本展开；图片与占位块用方括号标注
 * （宁可标注也不要静默丢内容）；分页符不产生文本行（否则会多出空行）。
 * 参数允许有洞（未取到的块是 `undefined`），下标即块下标。
 */
export function blocksToPlainText(blocks: Array<DocBlock | undefined>): string {
  const lines: string[] = [];
  for (const block of blocks) {
    if (!block) continue;
    const text = blockPlainText(block);
    if (text !== null) lines.push(text);
  }
  return lines.join("\n");
}

/* ------------------------- 分页 / 占位块合并计划 ------------------------- */

/** 这个块是否**强制开新页**（显式分页符 / 段前分页 / 分节符） */
export function startsNewPage(block: DocBlock): boolean {
  if (block.kind === "pageBreak") return true;
  if (block.kind === "paragraph") return block.pageBreakBefore || block.sectionBreak !== null;
  return false;
}

/**
 * 每个块的渲染计划（一遍扫完，供高度索引与渲染共用）：
 *  · `leader[i]`：块 i 所属「合并组」的首块下标（不合并时就是自己）；
 *  · `count[i]`：组的块数；**0 表示它是组内非首块（不渲染，高度记 0）**；
 *  · `startsPage[i]`：是否强制开新页（分页模式用）。
 *
 * 合并规则：**相邻**且 `kind === "unsupported"` 且 `label` 相同才算一组。
 * 中间夹任何别的块、或 label 不同（公式 / 文本框 / SmartArt 混排）都不合并 ——
 * 一份合同里 6 个「图形对象」卡片正是靠这条降噪成 1 张。
 * 未加载的块（`undefined`）会截断合并链：宁可少合并，也不要把不相邻的东西并到一起。
 */
export interface BlockPlan {
  leader: Int32Array;
  count: Int32Array;
  startsPage: Uint8Array;
}

export function planBlocks(total: number, blockAt: (index: number) => DocBlock | undefined): BlockPlan {
  const leader = new Int32Array(total);
  const count = new Int32Array(total);
  const startsPage = new Uint8Array(total);
  let index = 0;
  while (index < total) {
    const block = blockAt(index);
    startsPage[index] = block && startsNewPage(block) ? 1 : 0;
    if (block && block.kind === "unsupported") {
      let next = index + 1;
      while (next < total) {
        const candidate = blockAt(next);
        if (!candidate || candidate.kind !== "unsupported" || candidate.label !== block.label) break;
        startsPage[next] = startsNewPage(candidate) ? 1 : 0;
        next += 1;
      }
      leader[index] = index;
      count[index] = next - index;
      for (let member = index + 1; member < next; member += 1) {
        leader[member] = index;
        count[member] = 0;
      }
      index = next;
      continue;
    }
    leader[index] = index;
    count[index] = 1;
    index += 1;
  }
  return { leader, count, startsPage };
}

/** 大纲条目：按 `outlineLevel`（0..8）收集标题 */
export interface OutlineItem {
  /** 顶层块下标（点击后滚到这里） */
  index: number;
  /** 大纲级别（0 起） */
  level: number;
  /** 标题文字（取段落纯文本，去掉列表前缀那种噪音） */
  text: string;
}

/** 标记文字最大长度（侧栏里一行放不下，截断显示；tooltip 给全文） */
const OUTLINE_TEXT_LIMIT = 60;

/**
 * 收集大纲。**只看 `outlineLevel`**：`styleId` / `style` 在 WPS 与 Word 里五花八门
 * （`4`、`a9`、`heading 4`…），拿它判断标题必然漏；后端已把样式继承链算成大纲级别。
 * `outlineLevel` 为 null 的段落不进大纲。参数允许有洞（未取到的块是 `undefined`）。
 */
export function collectOutline(blocks: Array<DocBlock | undefined>): OutlineItem[] {
  const items: OutlineItem[] = [];
  blocks.forEach((block, index) => {
    if (!block || block.kind !== "paragraph" || block.outlineLevel === null) return;
    const level = Math.max(0, Math.min(8, Math.round(block.outlineLevel)));
    const body = block.text.replace(/[\n\t]/g, " ").trim();
    const prefix = block.list?.ordered && block.list.prefix ? `${block.list.prefix} ` : "";
    const raw = body ? `${prefix}${body}` : "";
    items.push({
      index,
      level,
      text: raw.length > OUTLINE_TEXT_LIMIT ? `${raw.slice(0, OUTLINE_TEXT_LIMIT)}…` : raw,
    });
  });
  return items;
}
