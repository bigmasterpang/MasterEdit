//! 表格「简单常用功能」：查找、结构操作、选区统计（0.20.0）。
//!
//! 定位是「简单的查看编辑」：整簿查找、行/列/工作表的增删改、选区数值统计。
//! **不做**公式引擎、样式编辑、图表等复杂内容。
//!
//! 本模块的命令：
//! - `spreadsheet_find` / `spreadsheet_stats`：只读，走 calamine（与 `office.rs` 同一套
//!   做法：企业透明加密文档先在**内存**里解密，绝不写明文临时文件）。两者读到的都是
//!   `office.rs::read_workbook_bytes` 的结果，因此**自动优先读内存影子**（见下）。
//! - `spreadsheet_structure` / `spreadsheet_state` / `spreadsheet_undo` / `spreadsheet_redo`
//!   / `spreadsheet_discard`：结构操作与它的撤销/重做/查询，全部围绕 `office_shadow` 的
//!   影子工作簿。
//!
//! ## 结构操作是「延迟写回 + 可撤销」
//! 0.20.0 起结构操作**不再直接落盘**：它只改内存里的影子工作簿（`office_shadow`），
//! 用户按 Ctrl+S（`spreadsheet_save`）才把影子的内容整簿写回文件。因此：
//! - 执行后磁盘字节一字不动，可以随时 `spreadsheet_discard` 放弃；
//! - 插入/删除/新建表的结果立刻能看见（所有读路径都优先读影子），但文件还没变；
//! - 每一次**结构改动**都会把「改动前」的整簿快照压进撤销栈，`spreadsheet_undo` /
//!   `spreadsheet_redo` 可以在内存里来回走（**保存也不清空撤销栈**，见 `office_shadow`）；
//! - 保存时才用 calamine 校验输出、才原子替换、才按原 4096 字节头重新加密企业加密文件；
//!   `.bak` 备份**默认关闭**（可显式要求）—— 全部复用 `office_write` 的写盘流水线。
//!
//! ## 结构操作的已知取舍（前端弹警告告知用户，这里同步写明）
//! - **超出已用范围 = 空操作**：插入/删除的 `at` 落在当前已用区域之外时**不报错**，直接
//!   成功返回，也不会扩张工作表尺寸（与 Excel 一致），并且**不留影子、不报 pending**
//!   （什么都没改，不该让文档变脏）。前端「在下方插入行 / 在右侧插入列」给出的正是
//!   「最后一行/列 + 1」，在小表或空表上点它不该被拦住。删除范围只有一部分在已用区域内
//!   时，按「能删多少删多少」处理，同样不报错。
//! - **公式不重算**：删除行/列会连带删掉该范围内的公式与样式（umya 行为）；剩下的公式
//!   只保留原有缓存值（可能过期），由 Excel / WPS 打开后自行计算。本模块**不会**自己
//!   去重算或调整公式引用，跨表引用（`=其它表!A1`）在插入/删除后可能指向别处 ——
//!   与 `office_write` 的「公式不重算」是同一条约定。
//! - **重命名工作表**不会改写其它工作表里对该表名的引用（`=旧表名!A1` 会变成 #REF!），
//!   需要用户自行检查。
//! - **原地写回**：结构操作没有「另存为」，保存时一律写回原文件（默认先备份 `<文件名>.bak`）。
//! - 插入/删除只影响目标工作表：**其它工作表分毫不动**（用的是 `Worksheet` 级 API，
//!   不是会遍历整簿的 `Workbook` 级 API）。
//! - 仍然会报错的情况：`count > 100000`、工作表名不存在、工作表重名、删除最后一张表、
//!   非 `.xlsx` 文件、同时有未保存结构改动的表格超过 8 个等。
//!
//! ## 复制工作表（`CopySheet`）的实现方式与性能边界
//! umya 3.1 没有「克隆工作表」的公开 API：新表必须由 `Workbook::new_sheet` 建出来才有
//! 正确的 sheetId 与默认格式（`set_sheet_id` 是 crate 私有，手工 `Worksheet::default()`
//! + `add_sheet` 会写出 `sheetId` 为空的工作表）。因此实现是**新建表 + 逐格复制**：
//! 用 `Cell::clone` 搬走源表每个已分配单元格（值 + 样式 + 超链接 + 公式），再复制列宽、
//! 行高与合并区域。成本与「已用区域」的格子数成正比，所以
//! `highest_column_and_row` 算出的格子数超过 `MAX_COPY_CELLS`（20 万）时直接报
//! 「表格过大，暂不支持复制」——这是本命令唯一的性能边界：20 万格以内是线性时间、
//! 内存开销约为该表数据量的一倍（克隆期间源表与目标表同时存在）。
//!
//! ## 查找与公式
//! 命中判定同时看**显示文本**与**公式文本**：`Data` 里存的是公式的缓存值，若只看值，
//! 搜 `SUM` 这种公式内容会一无所获；两者都算命中更实用。返回的 `text` 优先给显示值，
//! 只有公式命中时给公式文本（否则用户看不出为什么命中）。公式按 `=SUM(B2:B2)` 的形式
//! （与 `office.rs::spreadsheet_rows` 返回的 `f` 字段一致）参与匹配。

use crate::commands::image::civil_from_days;
// 读取与解析直接复用 office.rs 的实现（同一份内存解密 + calamine 解析 + 失败提示话术）
use crate::commands::office::{describe_parse_failure, parse_workbook, read_workbook_bytes, SheetMeta};
use crate::commands::office_shadow;
use crate::commands::office_write::{parse_number, CellEdit};
use calamine::{Data, Reader};
use std::borrow::Cow;
use std::path::Path;

/// 单次查找最多返回的命中数（达到上限即停止扫描）
const MAX_HITS: usize = 500;
/// 单条命中文本的字符上限（超出部分截断并加省略号）
const MAX_HIT_CHARS: usize = 200;
/// 单次结构操作允许的最大行/列数（防止前端传错把整簿拖死）
const MAX_OP_COUNT: u32 = 100_000;
/// 选区统计的单元格上限
const MAX_STATS_CELLS: u64 = 1_000_000;
/// 复制工作表的已用区域格子上限
const MAX_COPY_CELLS: u64 = 200_000;
/// xlsx 规格上限（Excel 2010+）：插入后不能越过
const MAX_ROWS: u32 = 1_048_576;
const MAX_COLS: u32 = 16_384;
/// 工作表名的字符上限（Excel 约定）
const SHEET_NAME_MAX_CHARS: usize = 31;

/* ------------------------------------------------------------------ */
/* 读取：企业加密内存解密 + calamine 解析 —— 与 office.rs 共用同一份实现  */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/* 单元格显示文本（口径与 office.rs::cell_text 一致）                    */
/* ------------------------------------------------------------------ */

/// 浮点显示：整数不带小数点，其余用最短往返表示（-0 归一化为 0）。
/// `pub(crate)`：公式求值的结果显示复用同一份口径（数字不带多余小数）。
pub(crate) fn format_float(f: f64) -> String {
    if !f.is_finite() {
        return format!("{f}");
    }
    let text = if f.fract() == 0.0 && f.abs() < 1e15 {
        format!("{f:.0}")
    } else {
        format!("{f}")
    };
    if text == "-0" {
        "0".to_string()
    } else {
        text
    }
}

/// 一天内的小数部分 → HH:MM:SS
fn format_time_of_day(frac: f64) -> String {
    let total = (frac * 86_400.0).round() as i64;
    let total = total.rem_euclid(86_400);
    format!(
        "{:02}:{:02}:{:02}",
        total / 3600,
        (total % 3600) / 60,
        total % 60
    )
}

/// Excel 序列日期 → 显示文本（1900 日期系统；见 office.rs::format_serial 的说明）
fn format_serial(serial: f64, is_duration: bool) -> String {
    if !serial.is_finite() {
        return format!("{serial}");
    }
    if is_duration {
        let total = (serial * 86_400.0).round() as i64;
        let sign = if total < 0 { "-" } else { "" };
        let t = total.abs();
        return format!("{sign}{}:{:02}:{:02}", t / 3600, (t % 3600) / 60, t % 60);
    }
    if serial < 0.0 {
        return format_float(serial);
    }
    if serial < 1.0 {
        return format_time_of_day(serial);
    }
    let days = serial.floor() as i64;
    let unix_days = if days >= 61 { days - 25569 } else { days - 25568 };
    let (y, m, d) = civil_from_days(unix_days);
    let frac = serial - serial.floor();
    if frac > 0.0 {
        format!("{y:04}-{m:02}-{d:02} {}", format_time_of_day(frac))
    } else {
        format!("{y:04}-{m:02}-{d:02}")
    }
}

/// 单元格显示文本：与前端网格里看到的一致（文本原样、数字紧凑、日期按序列号渲染）。
/// `pub(crate)`：公式求值把单元格转成值时复用同一份口径。
pub(crate) fn cell_display(data: &Data) -> String {
    match data {
        Data::Empty => String::new(),
        Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => s.clone(),
        Data::Int(i) => i.to_string(),
        Data::Float(f) => format_float(*f),
        Data::Bool(b) => (if *b { "true" } else { "false" }).to_string(),
        Data::DateTime(dt) => format_serial(dt.as_f64(), dt.is_duration()),
        Data::Error(_) => data.to_string(),
    }
}

/// 按字符边界截断（多字节字符不会被切开），截断时补省略号
fn truncate_chars(s: &str, max: usize) -> String {
    if s.chars().count() <= max {
        return s.to_string();
    }
    let mut out: String = s.chars().take(max).collect();
    out.push('…');
    out
}

/// calamine 的公式文本 → 与 `office.rs` 一致的形式（带前导 `=`，空串归一为 None）
fn formula_display(raw: &str) -> Option<String> {
    let text = raw.trim();
    if text.is_empty() {
        return None;
    }
    if text.starts_with('=') {
        Some(text.to_string())
    } else {
        Some(format!("={text}"))
    }
}

/* ------------------------------------------------------------------ */
/* 命令 1：查找                                                        */
/* ------------------------------------------------------------------ */

/// 查找选项（全部可选，缺省即最宽松的「不区分大小写、包含匹配、所有工作表」）
#[derive(serde::Deserialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct FindOptions {
    /// 区分大小写（默认 false）
    pub match_case: Option<bool>,
    /// 全字匹配（整格内容相等，忽略大小写时按大小写不敏感比较）
    pub whole_cell: Option<bool>,
    /// 只查某个工作表（None = 所有表）
    pub sheet: Option<String>,
}

/// 一处命中
#[derive(serde::Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct FindHit {
    pub sheet: String,
    /// 0 起的绝对行列号（与前端网格坐标一致）
    pub row: u32,
    pub col: u32,
    /// 命中单元格的显示文本（截断到 200 字符以内；仅公式命中时给公式文本）
    pub text: String,
}

/// 文本是否命中。`needle` 已按大小写规则归一化（不区分大小写时是小写形式）。
fn text_hit(text: &str, needle: &str, match_case: bool, whole_cell: bool) -> bool {
    let hay = if match_case {
        Cow::Borrowed(text)
    } else {
        Cow::Owned(text.to_lowercase())
    };
    if whole_cell {
        hay == needle
    } else {
        hay.contains(needle)
    }
}

/// 查找实现（同步，供命令在线程池里跑，也方便单元测试直接调用）
fn find_impl(
    path: &str,
    query: &str,
    options: Option<FindOptions>,
) -> Result<Vec<FindHit>, String> {
    // 空查询直接返回空数组（不去读文件：前端清空输入框时会频繁调用）
    if query.is_empty() {
        return Ok(Vec::new());
    }
    let options = options.unwrap_or_default();
    let match_case = options.match_case.unwrap_or(false);
    let whole_cell = options.whole_cell.unwrap_or(false);
    let needle = if match_case {
        query.to_string()
    } else {
        query.to_lowercase()
    };

    let (bytes, _encrypted) = read_workbook_bytes(Path::new(path))?;
    let mut workbook = parse_workbook(&bytes).map_err(|e| describe_parse_failure(&bytes, &e))?;

    // 工作表顺序 = 工作簿里的顺序，命中顺序即前端 F3 跳转顺序
    let names: Vec<String> = workbook.sheet_names();
    let targets: Vec<String> = match options
        .sheet
        .as_deref()
        .map(str::trim)
        .filter(|s| !s.is_empty())
    {
        Some(want) => {
            if !names.iter().any(|n| n == want) {
                return Err(format!("工作簿里没有工作表「{want}」"));
            }
            vec![want.to_string()]
        }
        None => names,
    };

    let mut hits: Vec<FindHit> = Vec::new();
    'sheets: for name in targets {
        let values = workbook
            .worksheet_range(&name)
            .map_err(|e| format!("读取工作表「{name}」失败: {e}"))?;
        // 公式为可选能力（xls / ods 可能不支持）：与 office.rs 一样失败即视为无公式
        let formulas = workbook.worksheet_formula(&name).ok();
        // Range 可能不从 A1 起算，因此一律用绝对坐标取值（与 office.rs::rows_impl 一致）
        let (Some(start), Some(end)) = (values.start(), values.end()) else {
            continue;
        };
        for row in start.0..=end.0 {
            for col in start.1..=end.1 {
                let Some(data) = values.get_value((row, col)) else {
                    continue;
                };
                if matches!(data, Data::Empty) {
                    continue;
                }
                let text = cell_display(data);
                let in_value = text_hit(&text, &needle, match_case, whole_cell);
                let formula = formulas
                    .as_ref()
                    .and_then(|range| range.get_value((row, col)))
                    .and_then(|raw| formula_display(raw));
                let in_formula = formula
                    .as_deref()
                    .map(|f| text_hit(f, &needle, match_case, whole_cell))
                    .unwrap_or(false);
                if !in_value && !in_formula {
                    continue;
                }
                // 优先显示单元格的值；只有公式命中时显示公式，用户才知道为什么命中
                let shown = if in_value {
                    text
                } else {
                    formula.unwrap_or_default()
                };
                hits.push(FindHit {
                    sheet: name.clone(),
                    row,
                    col,
                    text: truncate_chars(&shown, MAX_HIT_CHARS),
                });
                if hits.len() >= MAX_HITS {
                    break 'sheets;
                }
            }
        }
    }
    Ok(hits)
}

/// 在工作簿里查找文本（0.20.0「表格查找」）。
///
/// - 上限 500 条命中：**达到上限即停止扫描**，返回已找到的（顺序为工作表顺序 → 行 → 列，
///   前端按此顺序做 F3 跳转）；
/// - `query` 为空直接返回空数组（不读文件）；
/// - 企业透明加密文档在内存中解密后解析，不落明文；
/// - 公式单元格按公式文本（形如 `=SUM(B2:B2)`）**与**缓存值双向参与匹配；返回的 `text`
///   优先给显示值，只有公式命中时给公式文本。
#[tauri::command]
pub async fn spreadsheet_find(
    path: String,
    query: String,
    options: Option<FindOptions>,
) -> Result<Vec<FindHit>, String> {
    tauri::async_runtime::spawn_blocking(move || find_impl(&path, &query, options))
        .await
        .map_err(|e| format!("查找表格失败: {e}"))?
}

/* ------------------------------------------------------------------ */
/* 命令 2：结构操作                                                    */
/* ------------------------------------------------------------------ */

/// 一处结构变更。行/列坐标均为 **0 起**（与前端网格一致），Rust 侧 +1 后交给 umya。
#[derive(serde::Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum StructureOp {
    /// 在 `at` 行之前插入 `count` 行
    InsertRows { sheet: String, at: u32, count: u32 },
    /// 从 `at` 行开始删除 `count` 行（连带删掉范围内的公式与样式，不重算）
    DeleteRows { sheet: String, at: u32, count: u32 },
    /// 在 `at` 列之前插入 `count` 列
    InsertCols { sheet: String, at: u32, count: u32 },
    /// 从 `at` 列开始删除 `count` 列（连带删掉范围内的公式与样式，不重算）
    DeleteCols { sheet: String, at: u32, count: u32 },
    /// 新建工作表（`name` 为空时自动生成不与现有重名的 `SheetN`）
    AddSheet { name: Option<String> },
    /// 重命名工作表（重名直接报错）
    RenameSheet { sheet: String, name: String },
    /// 删除工作表（至少保留一张）
    DeleteSheet { sheet: String },
    /// 复制工作表（新建表 + 逐格复制值与样式；`name` 为空时自动生成）
    CopySheet {
        sheet: String,
        name: Option<String>,
    },
}

/// `count` 归一化：0 视为 1；过大直接报错（避免前端传错把整簿拖死）
fn normalize_count(count: u32, action: &str) -> Result<u32, String> {
    let count = if count == 0 { 1 } else { count };
    if count > MAX_OP_COUNT {
        return Err(format!(
            "{action}的行/列数过大：count={count}，单次最多 {MAX_OP_COUNT}"
        ));
    }
    Ok(count)
}

/// 工作簿里没有这张表时的统一报错
fn no_sheet(sheet: &str) -> String {
    format!("工作簿里没有工作表「{sheet}」")
}

/// 工作表名是否合法（Excel 约定：非空、≤31 字符、不含 `: \ / ? * [ ]`）
fn validate_sheet_name(name: &str) -> Result<(), String> {
    if name.is_empty() {
        return Err("工作表名不能为空".to_string());
    }
    if name.chars().count() > SHEET_NAME_MAX_CHARS {
        return Err(format!(
            "工作表名过长：「{name}」有 {} 个字符，最多 {SHEET_NAME_MAX_CHARS} 个",
            name.chars().count()
        ));
    }
    if let Some(bad) = name.chars().find(|c| matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']')) {
        return Err(format!("工作表名不能包含字符「{bad}」：{name}"));
    }
    Ok(())
}

/// 现有工作表里是否已有该名字
fn name_taken(workbook: &umya_spreadsheet::Workbook, name: &str) -> bool {
    workbook
        .sheet_collection_no_check()
        .iter()
        .any(|s| s.name() == name)
}

/// 自动生成工作表名：取第一个没被占用的 `SheetN`（N 从 1 起，避免与现有表重名）
fn next_sheet_name(workbook: &umya_spreadsheet::Workbook) -> String {
    for n in 1..=100_000u32 {
        let candidate = format!("Sheet{n}");
        if !name_taken(workbook, &candidate) {
            return candidate;
        }
    }
    // 10 万个 SheetN 全被占用的情形实际不存在，这里只是兜底（保证唯一且不 panic）
    let mut n = 100_001u32;
    while name_taken(workbook, &format!("Sheet{n}")) {
        n += 1;
    }
    format!("Sheet{n}")
}

/// 显式名字先校验，否则自动生成
fn resolve_sheet_name(
    workbook: &umya_spreadsheet::Workbook,
    name: Option<&str>,
) -> Result<String, String> {
    match name.map(str::trim).filter(|n| !n.is_empty()) {
        Some(explicit) => {
            validate_sheet_name(explicit)?;
            if name_taken(workbook, explicit) {
                return Err(format!("工作表「{explicit}」已存在"));
            }
            Ok(explicit.to_string())
        }
        None => Ok(next_sheet_name(workbook)),
    }
}

/// 行插入/删除点是否落在当前已用区域之外（0 起的 `at` 大于最后一个已用行号）。
///
/// 这种位置上的插入/删除等于**空操作**（与 Excel 一致）：直接成功返回，既不报错，
/// 也不扩张工作表尺寸（空表更是任何位置都算“之外”，插入不会凭空多出内容）。
/// 前端「在下方插入行」给出的正是「最后一行 + 1」，在小表 / 空表上点它不该被拦住，
/// 所以这里不做越界报错，只用来判断“是否真的改了东西”。
fn row_beyond_used(sheet: &umya_spreadsheet::Worksheet, at: u32) -> bool {
    let max_row = sheet.highest_column_and_row().1; // 1 起；空表为 0
    max_row == 0 || at > max_row - 1
}

/// 列版本，语义见 [`row_beyond_used`]
fn col_beyond_used(sheet: &umya_spreadsheet::Worksheet, at: u32) -> bool {
    let max_col = sheet.highest_column_and_row().0; // 1 起；空表为 0
    max_col == 0 || at > max_col - 1
}

/// 插入后不能越过 xlsx 规格上限（1_048_576 行 / 16_384 列）
fn check_insert_limit(at: u32, count: u32, limit: u32, unit: &str) -> Result<(), String> {
    if at as u64 + count as u64 > limit as u64 {
        return Err(format!(
            "插入{unit}超出 xlsx 规格上限：at={at} + count={count} 超过 {limit}"
        ));
    }
    Ok(())
}

/// 复制工作表：新建 + 逐格复制（详见模块注释里的说明与性能边界）
fn copy_sheet(
    workbook: &mut umya_spreadsheet::Workbook,
    source: &str,
    name: Option<&str>,
) -> Result<String, String> {
    let index = workbook
        .sheet_collection_no_check()
        .iter()
        .position(|s| s.name() == source)
        .ok_or_else(|| no_sheet(source))?;

    // 性能边界：按已用区域算格子数（复制成本与之成正比）
    let (cols, rows) = workbook.sheet_collection_no_check()[index].highest_column_and_row();
    let cells = cols as u64 * rows as u64;
    if cells > MAX_COPY_CELLS {
        return Err(format!(
            "表格过大，暂不支持复制：「{source}」已用区域约 {rows} 行 × {cols} 列（{cells} 格），超过 {MAX_COPY_CELLS} 格上限"
        ));
    }

    let new_name = resolve_sheet_name(workbook, name)?;
    workbook
        .new_sheet(new_name.clone())
        .map_err(|e| format!("新建工作表「{new_name}」失败: {e}"))?;

    // 新表刚被 push 到末尾：用 split_at_mut 同时拿到「只读的源表」和「可写的新表」
    let last = workbook.sheet_count() - 1;
    let sheets = workbook.sheet_collection_mut();
    let (head, tail) = sheets.split_at_mut(last);
    let dst = &mut tail[0];
    let src = &head[index];
    for cell in src.cells() {
        // Cell::clone 带上值、样式、超链接、公式；坐标不变，相对引用因此仍然有效
        dst.set_cell(cell.clone());
    }
    // 列宽 / 行高 / 合并区域一并复制，否则复制出来的表「瘦」得不像原表
    *dst.column_dimensions_mut() = src.column_dimensions().to_vec();
    *dst.row_dimensions_to_hashmap_mut() = src.row_dimensions_to_hashmap().clone();
    for merged in src.merge_cells() {
        dst.add_merge_cells(merged.range());
    }
    Ok(new_name)
}

/// 把一处结构变更应用到内存工作簿（失败即返回，磁盘文件尚未被碰过）
/// 把一处结构变更应用到工作簿（内存），返回**是否真的改了东西**。
///
/// `Ok(false)` 表示这次是空操作（`at` 落在已用范围之外）：`office_shadow` 据此决定
/// 不留影子、不报 pending。失败即返回 Err，此时调用方手里的工作簿只可能被前面的
/// `edits` 改过（结构改动一律“要么全做、要么没做”）。
pub(crate) fn apply_structure(
    workbook: &mut umya_spreadsheet::Workbook,
    op: &StructureOp,
) -> Result<bool, String> {
    match op {
        StructureOp::InsertRows { sheet, at, count } => {
            let count = normalize_count(*count, "插入")?;
            let target = workbook.sheet_by_name_mut(sheet).map_err(|_| no_sheet(sheet))?;
            // 超出已用范围的插入 = 空操作（成功返回、不扩张尺寸），见 row_beyond_used
            if row_beyond_used(target, *at) {
                return Ok(false);
            }
            check_insert_limit(*at, count, MAX_ROWS, "行")?;
            // umya 坐标 1 起：0 起的 at 表示「在第 at 行之前插入」
            target.insert_new_row(at + 1, count);
        }
        StructureOp::DeleteRows { sheet, at, count } => {
            let count = normalize_count(*count, "删除")?;
            let target = workbook.sheet_by_name_mut(sheet).map_err(|_| no_sheet(sheet))?;
            // 整段落在已用范围之外 = 空操作；部分重叠时 umya 只删掉范围内真实存在的行
            if row_beyond_used(target, *at) {
                return Ok(false);
            }
            target.remove_row(at + 1, count);
        }
        StructureOp::InsertCols { sheet, at, count } => {
            let count = normalize_count(*count, "插入")?;
            let target = workbook.sheet_by_name_mut(sheet).map_err(|_| no_sheet(sheet))?;
            if col_beyond_used(target, *at) {
                return Ok(false);
            }
            check_insert_limit(*at, count, MAX_COLS, "列")?;
            target.insert_new_column_by_index(at + 1, count);
        }
        StructureOp::DeleteCols { sheet, at, count } => {
            let count = normalize_count(*count, "删除")?;
            let target = workbook.sheet_by_name_mut(sheet).map_err(|_| no_sheet(sheet))?;
            if col_beyond_used(target, *at) {
                return Ok(false);
            }
            target.remove_column_by_index(at + 1, count);
        }
        StructureOp::AddSheet { name } => {
            let new_name = resolve_sheet_name(workbook, name.as_deref())?;
            workbook
                .new_sheet(new_name.clone())
                .map_err(|e| format!("新建工作表「{new_name}」失败: {e}"))?;
        }
        StructureOp::RenameSheet { sheet, name } => {
            let new_name = name.trim();
            validate_sheet_name(new_name)?;
            let index = workbook
                .sheet_collection_no_check()
                .iter()
                .position(|s| s.name() == sheet)
                .ok_or_else(|| no_sheet(sheet))?;
            if workbook.sheet_collection_no_check()[index].name() == new_name {
                return Ok(false); // 改成同一个名字：无事发生，不算重名错误
            }
            if name_taken(workbook, new_name) {
                return Err(format!("工作表「{new_name}」已存在"));
            }
            workbook
                .set_sheet_name(index, new_name)
                .map_err(|e| format!("重命名工作表失败: {e}"))?;
        }
        StructureOp::DeleteSheet { sheet } => {
            if !name_taken(workbook, sheet) {
                return Err(no_sheet(sheet));
            }
            if workbook.sheet_count() <= 1 {
                return Err("不能删除最后一个工作表：xlsx 至少保留一张工作表".to_string());
            }
            workbook
                .remove_sheet_by_name(sheet)
                .map_err(|e| format!("删除工作表「{sheet}」失败: {e}"))?;
        }
        StructureOp::CopySheet { sheet, name } => {
            copy_sheet(workbook, sheet, name.as_deref())?;
        }
    }
    Ok(true)
}

/// `spreadsheet_structure` 返回：操作后的工作表列表 + 未落盘状态
#[derive(serde::Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct StructureResult {
    pub path: String,
    /// 操作后的完整工作表列表（前端据此刷新底部标签，含新建/重命名/删除/复制的结果）
    pub sheets: Vec<SheetMeta>,
    /// 固定 false：结构操作不再直接落盘（保留字段便于前端与测试断言）
    pub saved: bool,
    /// 是否有未落盘的结构改动
    pub pending: bool,
}

/// `spreadsheet_state` 返回：未落盘状态 + 工作表列表
#[derive(serde::Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetState {
    pub path: String,
    pub sheets: Vec<SheetMeta>,
    pub pending: bool,
    /// 是否有可撤销的结构步骤（撤销栈非空）
    pub can_undo: bool,
    /// 是否有可重做的步骤（重做栈非空）
    pub can_redo: bool,
}

/// `spreadsheet_undo` / `spreadsheet_redo` 返回：撤销/重做之后的状态
#[derive(serde::Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct ShadowEditResult {
    pub path: String,
    /// 撤销/重做后的完整工作表列表
    pub sheets: Vec<SheetMeta>,
    /// 当前影子是否与磁盘不同
    pub pending: bool,
    pub can_undo: bool,
    pub can_redo: bool,
}

/// 结构操作实现（同步，供命令在线程池里跑，也方便单元测试直接调用）：
/// 只改内存影子，不落盘 —— 详见 `office_shadow` 模块注释。
pub(crate) fn structure_impl(
    store: &office_shadow::ShadowStore,
    path: &str,
    op: &StructureOp,
    edits: &[CellEdit],
) -> Result<StructureResult, String> {
    office_shadow::apply_op(store, path, op, edits)
}

/// 结构操作：插入/删除行列、增删改复制工作表（0.20.0「表格结构操作」）。
///
/// **延迟写回**：只改内存里的影子工作簿，磁盘一字不动；用户按 Ctrl+S
/// （`spreadsheet_save`）才把影子的内容真正写回文件，中间可以随时放弃
/// （`spreadsheet_discard`）。因此这里不再返回 `SaveResult`，而是返回操作后的
/// 工作表列表与 `pending` 状态。
///
/// - `edits` 是前端待提交的单元格编辑，**先应用编辑再做结构变更**（都只进影子）；
/// - `backup` 为兼容旧参数保留：结构操作不再直接落盘，备份只发生在保存那一刻；
/// - 行/列坐标 0 起；`count` 为 0 视为 1，超过 100000 报错；
/// - `at` 超出当前已用区域时不是错误，而是**空操作**（成功返回、不扩张工作表尺寸，
///   也不会留下影子）；删除范围只有一部分在已用区域内时「能删多少删多少」；
/// - 删除范围会连带删掉公式与样式，**不重算也不调整公式引用**（前端会弹警告）；
/// - 只有 `.xlsx` 可参与结构操作（xlsm 含宏、xls/xlsb/ods 结构不同）。
#[tauri::command]
pub async fn spreadsheet_structure(
    path: String,
    op: StructureOp,
    edits: Vec<CellEdit>,
    backup: Option<bool>,
) -> Result<StructureResult, String> {
    // 兼容旧参数：结构操作不再「每次操作都备份」，备份只在 spreadsheet_save 时发生
    let _ = backup;
    tauri::async_runtime::spawn_blocking(move || {
        structure_impl(&office_shadow::SHADOWS, &path, &op, &edits)
    })
    .await
    .map_err(|e| format!("表格结构操作失败: {e}"))?
}

/// 查询未落盘状态 + 工作表列表 + 可撤销/可重做标志
/// （前端在保存成功、外部改动、重新加载后调用）。
#[tauri::command]
pub async fn spreadsheet_state(path: String) -> Result<SpreadsheetState, String> {
    tauri::async_runtime::spawn_blocking(move || office_shadow::state(&office_shadow::SHADOWS, &path))
        .await
        .map_err(|e| format!("查询表格状态失败: {e}"))?
}

/// 撤销上一次表格改动（结构操作）。没有可撤销的返回 Err("没有可撤销的改动")。
///
/// 撤销只改内存影子：撤销之后影子重新变脏（`pending: true`），再按 Ctrl+S 才写回文件。
/// 保存**不会**清空撤销栈（用户明确要求「即使保存也不影响撤回」），所以刚保存完也能一路撤销。
#[tauri::command]
pub async fn spreadsheet_undo(path: String) -> Result<ShadowEditResult, String> {
    tauri::async_runtime::spawn_blocking(move || office_shadow::undo(&office_shadow::SHADOWS, &path))
        .await
        .map_err(|e| format!("撤销表格改动失败: {e}"))?
}

/// 重做上一次被撤销的改动（语义与 `spreadsheet_undo` 对称，栈空时报「没有可重做的改动」）。
#[tauri::command]
pub async fn spreadsheet_redo(path: String) -> Result<ShadowEditResult, String> {
    tauri::async_runtime::spawn_blocking(move || office_shadow::redo(&office_shadow::SHADOWS, &path))
        .await
        .map_err(|e| format!("重做表格改动失败: {e}"))?
}

/// 丢弃未落盘的结构改动（关闭文档、外部改动、用户主动放弃时调用）；返回是否真的丢弃了东西。
///
/// 丢弃后读到的又是磁盘上的内容（`pending` 归 false），**撤销/重做栈一并清空**。
#[tauri::command]
pub async fn spreadsheet_discard(path: String) -> Result<bool, String> {
    tauri::async_runtime::spawn_blocking(move || Ok(office_shadow::discard(&path)))
        .await
        .map_err(|e| format!("放弃表格改动失败: {e}"))?
}

/* ------------------------------------------------------------------ */
/* 命令 3：选区统计                                                    */
/* ------------------------------------------------------------------ */

/// 矩形选区（0 起、含端点，与前端网格一致）
#[derive(serde::Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct CellRange {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

/// 选区统计结果
#[derive(serde::Serialize, Debug, Default, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct RangeStats {
    /// 覆盖的单元格总数（含空；已按工作表已用范围裁剪）
    pub cells: usize,
    /// 非空单元格数
    pub non_empty: usize,
    /// 数值单元格数（含「纯数字文本」）
    pub numeric: usize,
    /// 数值求和（无数值时为 0）
    pub sum: f64,
    /// 数值平均（无数值时为 null）
    pub average: Option<f64>,
    /// 数值最小 / 最大（无数值时为 null）
    pub min: Option<f64>,
    pub max: Option<f64>,
}

/// 统计实现（同步，供命令在线程池里跑，也方便单元测试直接调用）
fn stats_impl(path: &str, sheet: &str, range: &CellRange) -> Result<RangeStats, String> {
    // 起止颠倒时交换（前端拖选方向不定）
    let (start_row, end_row) = if range.start_row <= range.end_row {
        (range.start_row, range.end_row)
    } else {
        (range.end_row, range.start_row)
    };
    let (start_col, end_col) = if range.start_col <= range.end_col {
        (range.start_col, range.end_col)
    } else {
        (range.end_col, range.start_col)
    };

    // 上限保护放在最前面：超大选区根本不去解析文件
    let area = (end_row - start_row + 1) as u64 * (end_col - start_col + 1) as u64;
    if area > MAX_STATS_CELLS {
        return Err(format!(
            "选区过大，暂不支持统计（选区 {area} 格，上限 {MAX_STATS_CELLS} 格）"
        ));
    }

    let (bytes, _encrypted) = read_workbook_bytes(Path::new(path))?;
    let mut workbook = parse_workbook(&bytes).map_err(|e| describe_parse_failure(&bytes, &e))?;
    let values = workbook
        .worksheet_range(sheet)
        .map_err(|_| no_sheet(sheet))?;

    // 超出工作表已用范围的部分忽略（不报错）：空表直接返回全零
    let Some((end_used_row, end_used_col)) = values.end() else {
        return Ok(RangeStats::default());
    };
    let last_row = end_row.min(end_used_row);
    let last_col = end_col.min(end_used_col);

    let mut stats = RangeStats::default();
    for row in start_row..=last_row {
        for col in start_col..=last_col {
            stats.cells += 1;
            // Range 可能不从 A1 起算：越界位置 get_value 返回 None，按空格处理
            let Some(data) = values.get_value((row, col)) else {
                continue;
            };
            if matches!(data, Data::Empty) {
                continue;
            }
            stats.non_empty += 1;
            // 数值判定：Int / Float，以及文本型的纯数字（CSV 导入后很常见）。
            // 布尔、日期、错误、空都不计入数值（与「能写回的数字」同一套语义）。
            let number = match data {
                Data::Int(i) => Some(*i as f64),
                Data::Float(f) => Some(*f),
                Data::String(s) => parse_number(s),
                _ => None,
            };
            if let Some(value) = number {
                stats.numeric += 1;
                stats.sum += value;
                stats.min = Some(stats.min.map_or(value, |m: f64| m.min(value)));
                stats.max = Some(stats.max.map_or(value, |m: f64| m.max(value)));
            }
        }
    }
    if stats.numeric > 0 {
        stats.average = Some(stats.sum / stats.numeric as f64);
    }
    Ok(stats)
}

/// 统计某个矩形区域的数值概况（求和/平均/计数/最大最小）。
///
/// - 区域坐标 0 起、**含端点**；`start > end` 时自动交换；
/// - 选区超过 1_000_000 格直接报错「选区过大，暂不支持统计」；
/// - 超出工作表已用范围的部分忽略，不报错；
/// - 数值 = `Int` / `Float` / 纯数字文本；布尔、日期、错误、空不计入数值
///   （但非空的都计入 `nonEmpty`）；
/// - 企业透明加密文档在内存中解密后统计，不落明文。
#[tauri::command]
pub async fn spreadsheet_stats(
    path: String,
    sheet: String,
    range: CellRange,
) -> Result<RangeStats, String> {
    tauri::async_runtime::spawn_blocking(move || stats_impl(&path, &sheet, &range))
        .await
        .map_err(|e| format!("统计选区失败: {e}"))?
}

#[cfg(test)]
mod tests {
    use super::*;
    // 仅测试用到：构造企业加密样本、用 calamine 独立读回校验写盘结果
    use crate::commands::esafenet;
    use crate::commands::office_shadow::{self, ShadowStore};
    use crate::commands::office_write::SaveResult;
    use calamine::{open_workbook_auto_from_rs, Sheets};
    use rust_xlsxwriter::{Color, Format, Workbook as XlsxWriter};
    use std::io::Cursor;
    use std::path::PathBuf;

    /// 每个测试一个独立目录，避免并行执行时互相干扰（也方便断言目录里没有残留文件）
    fn test_dir(name: &str) -> PathBuf {
        let dir = std::env::temp_dir()
            .join("masteredit-office-ops")
            .join(name);
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).expect("创建测试目录");
        dir
    }

    /// 样本：两张表。A1 带样式（加粗 + 黄底），含文本 / 数字 / 布尔 / 公式
    fn make_sample(path: &std::path::Path) {
        let mut wb = XlsxWriter::new();
        let styled = Format::new().set_bold().set_background_color(Color::Yellow);
        {
            let sheet = wb.add_worksheet();
            sheet.set_name("数据").unwrap();
            sheet.write_string_with_format(0, 0, "名称", &styled).unwrap();
            sheet.write_string(0, 1, "数量").unwrap();
            sheet.write_string(0, 2, "备注").unwrap();
            sheet.write_string(1, 0, "苹果").unwrap();
            sheet.write_number(1, 1, 12.5).unwrap();
            sheet.write_formula(1, 2, "=SUM(B2:B2)").unwrap();
            sheet.write_string(2, 0, "香蕉").unwrap();
            sheet.write_number(2, 1, 3.0).unwrap();
            sheet.write_string(3, 0, "合计").unwrap();
            sheet.write_formula(3, 1, "=SUM(B2:B3)").unwrap();
            sheet.write_string(4, 0, "Apple").unwrap();
            sheet.write_number(4, 1, 7.0).unwrap();
            sheet.write_string(4, 3, "x").unwrap();
        }
        {
            let second = wb.add_worksheet();
            second.set_name("第二表").unwrap();
            second.write_string(0, 0, "x").unwrap();
            second.write_number(1, 1, 7.0).unwrap();
        }
        wb.save(path).expect("生成样本 xlsx");
    }

    /// 用 calamine 重新打开**磁盘上的文件**
    fn open(path: &std::path::Path) -> Sheets<Cursor<Vec<u8>>> {
        let bytes = std::fs::read(path).expect("读取文件");
        open_workbook_auto_from_rs(Cursor::new(bytes)).expect("应能被 calamine 解析")
    }

    /// 从一份工作簿里读单元格（0 起绝对坐标）
    fn cell_of(wb: &mut Sheets<Cursor<Vec<u8>>>, sheet: &str, row: u32, col: u32) -> Data {
        wb.worksheet_range(sheet)
            .expect("工作表应存在")
            .get_value((row, col))
            .cloned()
            .unwrap_or(Data::Empty)
    }

    /// 从一份工作簿里读某表的已用行数
    fn rows_of_wb(wb: &mut Sheets<Cursor<Vec<u8>>>, sheet: &str) -> usize {
        let range = wb.worksheet_range(sheet).expect("工作表应存在");
        match range.end() {
            Some((row, _)) => row as usize + 1,
            None => 0,
        }
    }

    /// 从一份工作簿里读公式（去掉前导 =）
    fn formula_of_wb(
        wb: &mut Sheets<Cursor<Vec<u8>>>,
        sheet: &str,
        row: u32,
        col: u32,
    ) -> Option<String> {
        wb.worksheet_formula(sheet)
            .ok()?
            .get_value((row, col))
            .map(|s| s.trim().trim_start_matches('=').to_string())
            .filter(|s| !s.is_empty())
    }

    /// 用 umya 从字节里读单元格样式
    fn style_of_bytes(bytes: &[u8], sheet: &str, coord: (u32, u32)) -> umya_spreadsheet::Style {
        let wb = umya_spreadsheet::reader::xlsx::read_reader(Cursor::new(bytes), true).unwrap();
        wb.sheet_by_name(sheet)
            .unwrap()
            .cell(coord)
            .map(|c| c.style().clone())
            .unwrap_or_default()
    }

    fn edit(sheet: &str, row: u32, col: u32, kind: &str, value: &str) -> CellEdit {
        CellEdit {
            sheet: sheet.to_string(),
            row,
            col,
            kind: kind.to_string(),
            value: value.to_string(),
        }
    }

    fn insert_rows(sheet: &str, at: u32, count: u32) -> StructureOp {
        StructureOp::InsertRows {
            sheet: sheet.to_string(),
            at,
            count,
        }
    }

    fn delete_rows(sheet: &str, at: u32, count: u32) -> StructureOp {
        StructureOp::DeleteRows {
            sheet: sheet.to_string(),
            at,
            count,
        }
    }

    fn insert_cols(sheet: &str, at: u32, count: u32) -> StructureOp {
        StructureOp::InsertCols {
            sheet: sheet.to_string(),
            at,
            count,
        }
    }

    fn delete_cols(sheet: &str, at: u32, count: u32) -> StructureOp {
        StructureOp::DeleteCols {
            sheet: sheet.to_string(),
            at,
            count,
        }
    }

    /// 目录里的文件名（升序），用于断言「没有残留临时文件」
    fn dir_names(dir: &std::path::Path) -> Vec<String> {
        let mut names: Vec<String> = std::fs::read_dir(dir)
            .unwrap()
            .map(|e| e.unwrap().file_name().to_string_lossy().to_string())
            .collect();
        names.sort();
        names
    }

    /// 把 xlsx 包成企业加密文档（与 office_write 的加密测试同一套头）
    fn wrap_encrypted(plain: &[u8]) -> Vec<u8> {
        let mut header = vec![0u8; 4096];
        header[0..4].copy_from_slice(&[0xE0, 0xA8, 0x91, 0xE7]);
        header[8..12].copy_from_slice(&512u32.to_le_bytes());
        header[12..16].copy_from_slice(&4096u32.to_le_bytes());
        esafenet::encrypt_esafenet(&header, plain)
    }

    /// 测试夹具：一份样本 + 一个**独立**的影子仓库。
    ///
    /// 结构操作只改影子，所以断言必须分清两件事：
    /// - `cell/rows/sheets/formula/style` = 模拟前端读到的内容（有影子读影子）；
    /// - `disk_*` = 磁盘上的真实内容（结构操作后应保持原样，保存后才变）。
    struct Doc {
        dir: PathBuf,
        path: PathBuf,
        store: ShadowStore,
    }

    impl Doc {
        /// 标准样本（两张表：数据 / 第二表）
        fn sample(name: &str) -> Self {
            let doc = Self::empty(name, "sample.xlsx");
            make_sample(&doc.path);
            doc
        }

        /// 只有指定单元格的样本
        fn cells(name: &str, cells: &[(u32, u16, &str)]) -> Self {
            let doc = Self::empty(name, "sample.xlsx");
            make_sheet(&doc.path, "数据", cells);
            doc
        }

        fn empty(name: &str, file: &str) -> Self {
            let dir = test_dir(name);
            let path = dir.join(file);
            Doc {
                dir,
                path,
                store: ShadowStore::new(office_shadow::MAX_SHADOWS),
            }
        }

        fn structure(
            &self,
            op: &StructureOp,
            edits: &[CellEdit],
        ) -> Result<StructureResult, String> {
            office_shadow::apply_op(&self.store, self.path.to_str().unwrap(), op, edits)
        }

        fn state(&self) -> Result<SpreadsheetState, String> {
            office_shadow::state(&self.store, self.path.to_str().unwrap())
        }

        fn save(&self, edits: &[CellEdit], backup: bool) -> Result<SaveResult, String> {
            office_shadow::save(
                &self.store,
                self.path.to_str().unwrap(),
                None,
                edits,
                backup,
            )
        }

        fn discard(&self) -> bool {
            self.store.discard(self.path.to_str().unwrap())
        }

        fn backup_path(&self) -> PathBuf {
            let name = self.path.file_name().unwrap().to_string_lossy().to_string();
            self.dir.join(format!("{name}.bak"))
        }

        /// 磁盘字节
        fn disk_bytes(&self) -> Vec<u8> {
            std::fs::read(&self.path).expect("读取文件")
        }

        /// 模拟前端读到的内容（有影子就是影子内容）
        fn current(&self) -> Sheets<Cursor<Vec<u8>>> {
            let bytes = match self.store.bytes_of(self.path.to_str().unwrap()) {
                Some(shadow) => shadow.bytes,
                None => self.disk_bytes(),
            };
            open_workbook_auto_from_rs(Cursor::new(bytes)).expect("应能被 calamine 解析")
        }

        fn cell(&self, sheet: &str, row: u32, col: u32) -> Data {
            cell_of(&mut self.current(), sheet, row, col)
        }

        fn rows(&self, sheet: &str) -> usize {
            rows_of_wb(&mut self.current(), sheet)
        }

        fn sheets(&self) -> Vec<String> {
            self.current().sheet_names()
        }

        fn formula(&self, sheet: &str, row: u32, col: u32) -> Option<String> {
            formula_of_wb(&mut self.current(), sheet, row, col)
        }

        fn style(&self, sheet: &str, coord: (u32, u32)) -> umya_spreadsheet::Style {
            let bytes = match self.store.bytes_of(self.path.to_str().unwrap()) {
                Some(shadow) => shadow.bytes,
                None => self.disk_bytes(),
            };
            style_of_bytes(&bytes, sheet, coord)
        }

        /// 磁盘上的单元格（验证「还没保存，磁盘没变」）
        fn disk_cell(&self, sheet: &str, row: u32, col: u32) -> Data {
            cell_of(&mut open(&self.path), sheet, row, col)
        }

        fn disk_sheets(&self) -> Vec<String> {
            open(&self.path).sheet_names()
        }

        fn disk_rows(&self, sheet: &str) -> usize {
            rows_of_wb(&mut open(&self.path), sheet)
        }
    }

    /* --------------------------- 前后端数据契约 --------------------------- */

    /// 三条命令的入参 / 返回值 JSON 形状（前端按这些字段名调用，改坏了这里会先红）
    #[test]
    fn serde_wire_format_matches_frontend_contract() {
        // StructureOp：kind 标签 + camelCase 字段名
        let op: StructureOp =
            serde_json::from_str(r#"{"kind":"insertRows","sheet":"数据","at":1,"count":2}"#).unwrap();
        match op {
            StructureOp::InsertRows { sheet, at, count } => {
                assert_eq!((sheet.as_str(), at, count), ("数据", 1, 2));
            }
            other => panic!("反序列化成了 {other:?}"),
        }
        let op: StructureOp = serde_json::from_str(r#"{"kind":"addSheet","name":null}"#).unwrap();
        assert!(matches!(op, StructureOp::AddSheet { name: None }));
        let op: StructureOp =
            serde_json::from_str(r#"{"kind":"copySheet","sheet":"数据","name":"副本"}"#).unwrap();
        assert!(matches!(op, StructureOp::CopySheet { name: Some(_), .. }));
        let op: StructureOp = serde_json::from_str(r#"{"kind":"deleteSheet","sheet":"数据"}"#).unwrap();
        assert!(matches!(op, StructureOp::DeleteSheet { .. }));
        // 未知 kind 必须报错，而不是静默变成某个操作
        assert!(serde_json::from_str::<StructureOp>(r#"{"kind":"nope"}"#).is_err());

        // FindOptions 入参
        let opts: FindOptions = serde_json::from_str(
            r#"{"matchCase":true,"wholeCell":false,"sheet":"数据"}"#,
        )
        .unwrap();
        assert_eq!(opts.match_case, Some(true));
        assert_eq!(opts.whole_cell, Some(false));
        assert_eq!(opts.sheet.as_deref(), Some("数据"));
        // 三个字段都可省略
        let opts: FindOptions = serde_json::from_str("{}").unwrap();
        assert!(opts.match_case.is_none() && opts.sheet.is_none());

        // CellRange 入参
        let range: CellRange =
            serde_json::from_str(r#"{"startRow":0,"startCol":1,"endRow":2,"endCol":3}"#).unwrap();
        assert_eq!((range.start_row, range.start_col, range.end_row, range.end_col), (0, 1, 2, 3));

        // 返回值：FindHit / RangeStats 都是 camelCase
        let hit = FindHit {
            sheet: "数据".to_string(),
            row: 1,
            col: 2,
            text: "x".to_string(),
        };
        assert_eq!(
            serde_json::to_string(&hit).unwrap(),
            r#"{"sheet":"数据","row":1,"col":2,"text":"x"}"#
        );
        let stats = RangeStats {
            cells: 6,
            non_empty: 4,
            numeric: 2,
            sum: 3.5,
            average: Some(1.75),
            min: Some(1.0),
            max: Some(2.5),
        };
        let json = serde_json::to_string(&stats).unwrap();
        assert!(json.contains(r#""cells":6"#), "{json}");
        assert!(json.contains(r#""nonEmpty":4"#), "{json}");
        assert!(json.contains(r#""numeric":2"#), "{json}");
        assert!(json.contains(r#""average":1.75"#), "{json}");
        assert!(json.contains(r#""min":1.0"#), "{json}");
        // 无数值时是 null，而不是缺字段（前端直接读 .average 判空）
        let json = serde_json::to_string(&RangeStats::default()).unwrap();
        assert!(json.contains(r#""average":null"#), "{json}");
        assert!(json.contains(r#""min":null"#) && json.contains(r#""max":null"#), "{json}");

        // SaveResult 沿用 office_write 的形状
        let saved = SaveResult {
            path: "C:\\a.xlsx".to_string(),
            size: 10,
            modified_at: 20,
            saved_cells: 3,
            backup_path: None,
        };
        let json = serde_json::to_string(&saved).unwrap();
        assert!(json.contains(r#""modifiedAt":20"#), "{json}");
        assert!(json.contains(r#""savedCells":3"#), "{json}");
        assert!(json.contains(r#""backupPath":null"#), "{json}");

        // 结构操作（延迟写回）：StructureResult / SpreadsheetState 的字段名与形状
        let result = StructureResult {
            path: "C:\\a.xlsx".to_string(),
            sheets: vec![SheetMeta {
                name: "数据".to_string(),
            }],
            saved: false,
            pending: true,
        };
        let json = serde_json::to_string(&result).unwrap();
        assert_eq!(
            json,
            r#"{"path":"C:\\a.xlsx","sheets":[{"name":"数据"}],"saved":false,"pending":true}"#
        );
        let state = serde_json::to_string(&SpreadsheetState {
            path: "C:\\a.xlsx".to_string(),
            sheets: vec![],
            pending: false,
            can_undo: true,
            can_redo: false,
        })
        .unwrap();
        assert_eq!(
            state,
            r#"{"path":"C:\\a.xlsx","sheets":[],"pending":false,"canUndo":true,"canRedo":false}"#
        );

        // 撤销/重做结果
        let edited = serde_json::to_string(&ShadowEditResult {
            path: "C:\\a.xlsx".to_string(),
            sheets: vec![SheetMeta {
                name: "数据".to_string(),
            }],
            pending: true,
            can_undo: false,
            can_redo: true,
        })
        .unwrap();
        assert_eq!(
            edited,
            r#"{"path":"C:\\a.xlsx","sheets":[{"name":"数据"}],"pending":true,"canUndo":false,"canRedo":true}"#
        );
    }

    /* ------------------------------ 查找 ------------------------------ */

    /// 命中顺序：工作表顺序 → 行 → 列；多个表、多行多列都能命中
    #[test]
    fn find_returns_hits_in_row_major_order_across_sheets() {
        let dir = test_dir("find-order");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let p = path.to_str().unwrap();

        let hits = find_impl(p, "苹果", None).expect("查找应成功");
        assert_eq!(hits.len(), 1);
        assert_eq!(
            (hits[0].sheet.as_str(), hits[0].row, hits[0].col),
            ("数据", 1, 0)
        );
        assert_eq!(hits[0].text, "苹果");

        // 「x」在 数据!D5 与 第二表!A1 各一处：顺序按表顺序、行序、列序
        let hits = find_impl(p, "x", None).expect("查找应成功");
        let places: Vec<(&str, u32, u32)> = hits
            .iter()
            .map(|h| (h.sheet.as_str(), h.row, h.col))
            .collect();
        assert_eq!(places, vec![("数据", 4, 3), ("第二表", 0, 0)]);
        assert_eq!(hits[1].text, "x");

        // 数字文本参与匹配，且按显示文本给出（12.5 不是 12.500000）
        let hits = find_impl(p, "12.5", None).unwrap();
        assert_eq!((hits[0].sheet.as_str(), hits[0].row, hits[0].col), ("数据", 1, 1));
        assert_eq!(hits[0].text, "12.5");

        // 跨表：第二表的 7 与 数据的 7 都能命中，顺序按表顺序
        let hits = find_impl(p, "7", None).unwrap();
        let places: Vec<(&str, u32, u32)> = hits
            .iter()
            .map(|h| (h.sheet.as_str(), h.row, h.col))
            .collect();
        assert_eq!(places, vec![("数据", 4, 1), ("第二表", 1, 1)]);
    }

    /// match_case / whole_cell 的行为
    #[test]
    fn find_respects_match_case_and_whole_cell() {
        let dir = test_dir("find-case");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let p = path.to_str().unwrap();

        let opts = |case: bool, whole: bool| {
            Some(FindOptions {
                match_case: Some(case),
                whole_cell: Some(whole),
                sheet: None,
            })
        };

        // 不区分大小写：apple 命中 A5 的 "Apple"
        let hits = find_impl(p, "apple", opts(false, false)).unwrap();
        assert_eq!(hits.len(), 1);
        assert_eq!((hits[0].row, hits[0].col), (4, 0));

        // 区分大小写：大小写不符则不命中
        assert!(find_impl(p, "apple", opts(true, false)).unwrap().is_empty());

        // 全字匹配：「苹」是 A2「苹果」的一部分，全字匹配不命中，包含匹配命中
        assert_eq!(find_impl(p, "苹", opts(false, false)).unwrap().len(), 1);
        assert!(find_impl(p, "苹", opts(false, true)).unwrap().is_empty());

        // 全字匹配 + 忽略大小写：整格 "Apple" vs 查询 "apple" 命中
        let hits = find_impl(p, "apple", opts(false, true)).unwrap();
        assert_eq!(hits.len(), 1);
        assert_eq!(hits[0].text, "Apple");
    }

    /// 限定工作表：只查指定表；表名不存在要明确报错
    #[test]
    fn find_can_be_limited_to_one_sheet() {
        let dir = test_dir("find-sheet");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let p = path.to_str().unwrap();

        // 「x」在 数据 与 第二表 各一处，限定后只剩第二表那条
        let hits = find_impl(
            p,
            "x",
            Some(FindOptions {
                sheet: Some("第二表".to_string()),
                ..Default::default()
            }),
        )
        .unwrap();
        assert_eq!(hits.len(), 1);
        assert_eq!((hits[0].sheet.as_str(), hits[0].row, hits[0].col), ("第二表", 0, 0));

        // 第二表 B2 = 7，数据 B5 = 7：限定后只剩第二表那条
        let hits = find_impl(
            p,
            "7",
            Some(FindOptions {
                sheet: Some("第二表".to_string()),
                ..Default::default()
            }),
        )
        .unwrap();
        assert_eq!(hits.len(), 1);
        assert_eq!(hits[0].sheet, "第二表");

        let err = find_impl(
            p,
            "7",
            Some(FindOptions {
                sheet: Some("不存在".to_string()),
                ..Default::default()
            }),
        )
        .expect_err("表名不存在应报错");
        assert!(err.contains("不存在"), "错误信息应含表名：{err}");
    }

    /// 空查询直接返回空数组（即使文件不存在也不报错）
    #[test]
    fn find_empty_query_returns_empty_without_reading_file() {
        let hits = find_impl(r"C:\definitely\not\here.xlsx", "", None).expect("空查询不应报错");
        assert!(hits.is_empty());
    }

    /// 公式单元格按公式文本参与匹配（返回的 text 是公式本身）
    #[test]
    fn find_matches_formula_text() {
        let dir = test_dir("find-formula");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let p = path.to_str().unwrap();

        let hits = find_impl(p, "SUM(B2:B2)", None).unwrap();
        assert_eq!(hits.len(), 1, "应命中 C2 的公式：{hits:?}");
        assert_eq!((hits[0].row, hits[0].col), (1, 2));
        assert_eq!(hits[0].text, "=SUM(B2:B2)", "只有公式命中时 text 应是公式文本");

        // 忽略大小写同样命中公式
        let hits = find_impl(p, "sum(b2:b2)", None).unwrap();
        assert_eq!(hits.len(), 1);

        // 两个公式里都有 SUM：默认（不区分大小写）命中两条，区分大小写时 "sum" 一条都不命中
        let hits = find_impl(p, "sum", None).unwrap();
        assert_eq!(hits.len(), 2, "C2 与 B4 的公式都应命中：{hits:?}");
        let places: Vec<(u32, u32)> = hits.iter().map(|h| (h.row, h.col)).collect();
        assert_eq!(places, vec![(1, 2), (3, 1)], "行序：C2 在 B4 之前");
        assert!(find_impl(
            p,
            "sum",
            Some(FindOptions {
                match_case: Some(true),
                ..Default::default()
            })
        )
        .unwrap()
        .is_empty());
    }

    /// 500 条上限：造 600 个命中，只返回前 500 条且顺序正确
    #[test]
    fn find_stops_at_five_hundred_hits() {
        let dir = test_dir("find-cap");
        let path = dir.join("big.xlsx");
        {
            let mut wb = XlsxWriter::new();
            let sheet = wb.add_worksheet();
            sheet.set_name("命中").unwrap();
            for row in 0..600u32 {
                sheet.write_string(row, 0, "hit").unwrap();
            }
            wb.save(&path).unwrap();
        }
        let hits = find_impl(path.to_str().unwrap(), "hit", None).unwrap();
        assert_eq!(hits.len(), MAX_HITS, "应在上限处停止扫描");
        // 顺序是行优先，第 500 条应是第 500 行（0 起 499）
        assert_eq!((hits[0].row, hits[0].col), (0, 0));
        assert_eq!((hits[MAX_HITS - 1].row, hits[MAX_HITS - 1].col), (499, 0));
    }

    /// 超长命中文本按字符边界截断到 200 字符（+ 省略号）
    #[test]
    fn find_truncates_long_hit_text() {
        let dir = test_dir("find-truncate");
        let path = dir.join("long.xlsx");
        let long = format!("{}尾巴", "中".repeat(400));
        {
            let mut wb = XlsxWriter::new();
            let sheet = wb.add_worksheet();
            sheet.write_string(0, 0, &long).unwrap();
            wb.save(&path).unwrap();
        }
        let hits = find_impl(path.to_str().unwrap(), "尾巴", None).unwrap();
        assert_eq!(hits.len(), 1);
        let text = &hits[0].text;
        assert_eq!(text.chars().count(), MAX_HIT_CHARS + 1, "应是 200 字符 + 省略号");
        assert!(text.ends_with('…'));
        assert!(text.starts_with("中中"));
    }

    /* --------------------------- 结构操作：行 --------------------------- */

    /// 插入行：影子里的内容下移、总行数正确、另一张表不受影响、**磁盘一字不动**
    #[test]
    fn insert_rows_shifts_content_and_keeps_other_sheets() {
        let doc = Doc::sample("insert-rows");
        let before_rows = doc.rows("数据");
        let disk = doc.disk_bytes();

        let result = doc
            .structure(&insert_rows("数据", 1, 2), &[])
            .expect("插入行应成功");
        assert!(!result.saved, "结构操作不再直接落盘");
        assert!(result.pending, "有未落盘的结构改动");
        assert_eq!(result.path, doc.path.to_string_lossy());
        assert_eq!(
            result
                .sheets
                .iter()
                .map(|s| s.name.as_str())
                .collect::<Vec<_>>(),
            vec!["数据", "第二表"]
        );

        // 影子内容：第 1 行（0 起）的「苹果 / 12.5」下移到第 3 行
        assert_eq!(doc.cell("数据", 3, 0), Data::String("苹果".into()));
        assert_eq!(doc.cell("数据", 3, 1), Data::Float(12.5));
        assert_eq!(doc.cell("数据", 0, 0), Data::String("名称".into()));
        assert_eq!(doc.cell("数据", 1, 0), Data::Empty, "插入的两行应为空");
        assert_eq!(doc.rows("数据"), before_rows + 2, "总行数应 +2");
        // 其它工作表分毫不动
        assert_eq!(doc.cell("第二表", 0, 0), Data::String("x".into()));
        assert_eq!(doc.cell("第二表", 1, 1), Data::Float(7.0));

        // 磁盘完全没变
        assert_eq!(doc.disk_bytes(), disk, "结构操作不该改磁盘");
        assert_eq!(doc.disk_rows("数据"), before_rows);
    }

    /// 删除行：影子里的后续行上移、总行数正确、另一张表不受影响
    #[test]
    fn delete_rows_pulls_content_up_and_keeps_other_sheets() {
        let doc = Doc::sample("delete-rows");
        let before_rows = doc.rows("数据");
        let disk = doc.disk_bytes();

        doc.structure(&delete_rows("数据", 1, 1), &[])
            .expect("删除行应成功");

        assert_eq!(doc.cell("数据", 1, 0), Data::String("香蕉".into()));
        assert_eq!(doc.cell("数据", 1, 1), Data::Float(3.0));
        assert_eq!(doc.cell("数据", 2, 0), Data::String("合计".into()));
        assert_eq!(doc.rows("数据"), before_rows - 1);
        assert_eq!(doc.cell("第二表", 1, 1), Data::Float(7.0));
        assert_eq!(doc.disk_bytes(), disk);
    }

    /// 删除行会连带删掉范围内的公式（umya 行为，前端弹警告）
    #[test]
    fn delete_rows_removes_formulas_in_range() {
        let doc = Doc::sample("delete-formula");

        doc.structure(&delete_rows("数据", 3, 1), &[])
            .expect("删除应成功");

        // B4 的 =SUM(B2:B3) 随行被删掉，上移过来的「Apple」顶到第 3 行
        assert_eq!(doc.formula("数据", 3, 1), None, "被删行的公式不应残留");
        assert_eq!(doc.cell("数据", 3, 0), Data::String("Apple".into()));
        // 未被删的公式仍在（C2 的 =SUM(B2:B2)）
        assert_eq!(doc.formula("数据", 1, 2).as_deref(), Some("SUM(B2:B2)"));
    }

    /* --------------------------- 结构操作：列 --------------------------- */

    /// 插入列：影子里的原第 N 列右移、另一张表不受影响
    #[test]
    fn insert_cols_shifts_content() {
        let doc = Doc::sample("insert-cols");
        let disk = doc.disk_bytes();

        doc.structure(&insert_cols("数据", 1, 1), &[])
            .expect("插入列应成功");

        assert_eq!(doc.cell("数据", 1, 0), Data::String("苹果".into()));
        assert_eq!(doc.cell("数据", 1, 2), Data::Float(12.5), "原 B 列右移到 C 列");
        assert_eq!(doc.cell("数据", 1, 1), Data::Empty, "新列应为空");
        assert_eq!(doc.cell("第二表", 1, 1), Data::Float(7.0));
        assert_eq!(doc.disk_bytes(), disk);
    }

    /// 删除列：影子里的后续列左移
    #[test]
    fn delete_cols_pulls_content_left() {
        let doc = Doc::sample("delete-cols");
        let disk = doc.disk_bytes();

        doc.structure(&delete_cols("数据", 1, 1), &[])
            .expect("删除列应成功");

        // 原 B 列（数量）被删，原 C 列（公式）左移到 B 列
        assert!(
            doc.formula("数据", 1, 1).is_some(),
            "原 C 列的公式应左移到 B 列"
        );
        assert_eq!(doc.cell("数据", 1, 0), Data::String("苹果".into()));
        assert_eq!(doc.cell("第二表", 1, 1), Data::Float(7.0));
        assert_eq!(doc.disk_bytes(), disk);
    }

    /* ------------------- 编辑 + 结构：累积到保存那一刻 ------------------- */

    /// edits 与结构操作一起提交：都先进影子，保存时才一起落盘
    #[test]
    fn edits_and_structure_accumulate_until_save() {
        let doc = Doc::sample("edit-and-structure");
        let disk = doc.disk_bytes();

        let edits = [
            edit("数据", 0, 0, "text", "标题"),
            edit("数据", 0, 1, "number", "99"),
        ];
        let result = doc
            .structure(&insert_rows("数据", 1, 1), &edits)
            .expect("应成功");
        assert!(result.pending);

        // 影子：编辑生效 + 结构变更生效
        assert_eq!(doc.cell("数据", 0, 0), Data::String("标题".into()));
        assert_eq!(doc.cell("数据", 0, 1), Data::Float(99.0));
        assert_eq!(doc.cell("数据", 2, 0), Data::String("苹果".into()));
        // 磁盘：既没有编辑也没有结构变更
        assert_eq!(doc.disk_bytes(), disk);
        assert_eq!(doc.disk_cell("数据", 0, 0), Data::String("名称".into()));
        assert_eq!(doc.disk_cell("数据", 1, 0), Data::String("苹果".into()));

        // 保存：两者一起落盘，.bak 是改动前的磁盘字节
        let saved = doc.save(&edits, true).expect("保存应成功");
        assert_eq!(saved.saved_cells, 2, "应记录 2 处编辑");
        assert_eq!(doc.disk_cell("数据", 0, 0), Data::String("标题".into()));
        assert_eq!(doc.disk_cell("数据", 2, 0), Data::String("苹果".into()));
        assert_eq!(std::fs::read(doc.backup_path()).unwrap(), disk);
        assert!(!doc.store.is_pending(doc.path.to_str().unwrap()));
    }

    /// 编辑本身非法（数字解析失败）时：不建立影子、磁盘也不动
    #[test]
    fn invalid_edit_aborts_the_whole_structure_operation() {
        let doc = Doc::sample("bad-edit");
        let disk = doc.disk_bytes();

        let edits = [edit("数据", 0, 0, "number", "不是数字")];
        let err = doc
            .structure(&insert_rows("数据", 1, 1), &edits)
            .expect_err("非法编辑应报错");
        assert!(err.contains("不是数字"), "实际提示：{err}");
        assert_eq!(doc.disk_bytes(), disk, "失败时原文件不得改变");
        assert!(
            !doc.store.has(doc.path.to_str().unwrap()),
            "失败不该留下影子"
        );
        assert_eq!(dir_names(&doc.dir), vec!["sample.xlsx".to_string()]);
    }

    /* --------------------------- 工作表级操作 --------------------------- */

    /// 新建工作表：显式名字、自动名字（不与现有重名）、重名报错、非法名字报错；磁盘不动
    #[test]
    fn add_sheet_names_and_validation() {
        let doc = Doc::sample("add-sheet");
        let disk = doc.disk_bytes();

        let result = doc
            .structure(
                &StructureOp::AddSheet {
                    name: Some("报表".to_string()),
                },
                &[],
            )
            .expect("新建表应成功");
        assert!(result.pending);
        assert_eq!(doc.sheets(), vec!["数据", "第二表", "报表"]);
        assert_eq!(doc.disk_sheets(), vec!["数据", "第二表"], "磁盘还没变");

        // 名字为空/None → 自动生成，且不与现有重名
        doc.structure(&StructureOp::AddSheet { name: None }, &[])
            .unwrap();
        let names = doc.sheets();
        let generated = names.last().unwrap().clone();
        assert_eq!(generated, "Sheet1", "应生成第一个没被占用的 SheetN");
        assert_eq!(names.iter().filter(|n| *n == &generated).count(), 1);

        // 重名报错
        let err = doc
            .structure(
                &StructureOp::AddSheet {
                    name: Some("数据".to_string()),
                },
                &[],
            )
            .expect_err("重名应报错");
        assert!(err.contains("已存在"), "实际提示：{err}");

        // 非法字符 / 过长
        let err = doc
            .structure(
                &StructureOp::AddSheet {
                    name: Some("坏[名]".to_string()),
                },
                &[],
            )
            .expect_err("非法名字应报错");
        assert!(err.contains('['), "实际提示：{err}");
        let err = doc
            .structure(
                &StructureOp::AddSheet {
                    name: Some("长".repeat(40)),
                },
                &[],
            )
            .expect_err("过长名字应报错");
        assert!(err.contains("过长"), "实际提示：{err}");

        // 一路下来磁盘都没有变
        assert_eq!(doc.disk_bytes(), disk);
        // 影子里的表列表就是前端要刷新的那份
        assert_eq!(
            doc.state().unwrap().sheets.len(),
            4,
            "数据 / 第二表 / 报表 / Sheet1"
        );
        assert!(doc.state().unwrap().pending);
    }

    /// 重命名：名字与内容都对；重名报错；改成自己不算错
    #[test]
    fn rename_sheet_checks_duplicates() {
        let doc = Doc::sample("rename-sheet");

        doc.structure(
            &StructureOp::RenameSheet {
                sheet: "第二表".to_string(),
                name: "新名字".to_string(),
            },
            &[],
        )
        .expect("重命名应成功");
        assert_eq!(doc.sheets(), vec!["数据", "新名字"], "顺序应保持");
        assert_eq!(doc.cell("新名字", 1, 1), Data::Float(7.0));

        let err = doc
            .structure(
                &StructureOp::RenameSheet {
                    sheet: "新名字".to_string(),
                    name: "数据".to_string(),
                },
                &[],
            )
            .expect_err("重名应报错");
        assert!(err.contains("已存在"), "实际提示：{err}");

        // 改成自己：无操作，不报错
        doc.structure(
            &StructureOp::RenameSheet {
                sheet: "新名字".to_string(),
                name: "新名字".to_string(),
            },
            &[],
        )
        .expect("改成自己应视为无操作");

        // 表名不存在
        let err = doc
            .structure(
                &StructureOp::RenameSheet {
                    sheet: "没有这张表".to_string(),
                    name: "x".to_string(),
                },
                &[],
            )
            .expect_err("表不存在应报错");
        assert!(err.contains("没有这张表"), "实际提示：{err}");
    }

    /// 删除工作表；删最后一张要报错
    #[test]
    fn delete_sheet_keeps_at_least_one() {
        let doc = Doc::sample("delete-sheet");

        doc.structure(
            &StructureOp::DeleteSheet {
                sheet: "第二表".to_string(),
            },
            &[],
        )
        .expect("删除工作表应成功");
        assert_eq!(doc.sheets(), vec!["数据"]);
        assert_eq!(doc.disk_sheets(), vec!["数据", "第二表"]);

        let err = doc
            .structure(
                &StructureOp::DeleteSheet {
                    sheet: "数据".to_string(),
                },
                &[],
            )
            .expect_err("删最后一张表应报错");
        assert!(err.contains("最后一个工作表"), "实际提示：{err}");
        assert_eq!(doc.sheets(), vec!["数据"]);

        // 表不存在同样报错
        let err = doc
            .structure(
                &StructureOp::DeleteSheet {
                    sheet: "没有".to_string(),
                },
                &[],
            )
            .expect_err("表不存在应报错");
        assert!(err.contains("没有"), "实际提示：{err}");
    }

    /// 复制工作表：值、样式、列宽都跟着走；自动命名不与现有重名
    #[test]
    fn copy_sheet_copies_values_and_styles() {
        let doc = Doc::sample("copy-sheet");

        doc.structure(
            &StructureOp::CopySheet {
                sheet: "数据".to_string(),
                name: Some("副本".to_string()),
            },
            &[],
        )
        .expect("复制工作表应成功");
        assert_eq!(doc.sheets(), vec!["数据", "第二表", "副本"]);

        // 值一致
        for (row, col) in [(0u32, 0u32), (1, 0), (1, 1), (2, 0), (2, 1), (4, 0), (4, 3)] {
            assert_eq!(
                doc.cell("副本", row, col),
                doc.cell("数据", row, col),
                "({row},{col}) 的值应与源表一致"
            );
        }
        // 样式一致（A1 是加粗 + 黄底）
        assert_eq!(
            doc.style("副本", (1, 1)),
            doc.style("数据", (1, 1)),
            "复制后的样式应与源表一致"
        );
        // 公式也带过来
        assert_eq!(doc.formula("副本", 1, 2).as_deref(), Some("SUM(B2:B2)"));

        // 自动命名
        doc.structure(
            &StructureOp::CopySheet {
                sheet: "数据".to_string(),
                name: None,
            },
            &[],
        )
        .unwrap();
        assert_eq!(doc.sheets().last().unwrap(), "Sheet1");

        // 源表不存在
        let err = doc
            .structure(
                &StructureOp::CopySheet {
                    sheet: "没有这张表".to_string(),
                    name: None,
                },
                &[],
            )
            .expect_err("源表不存在应报错");
        assert!(err.contains("没有这张表"), "实际提示：{err}");
    }

    /// 复制过大的表：返回「表格过大，暂不支持复制」
    #[test]
    fn copy_sheet_rejects_huge_sheet() {
        let doc = Doc::empty("copy-huge", "huge.xlsx");
        // 用 umya 造一张「已用区域 500×500 = 25 万格」但只有一个值的表
        {
            let mut wb = umya_spreadsheet::new_file();
            let sheet = wb.sheet_by_name_mut("Sheet1").unwrap();
            sheet.set_name("大表");
            sheet.cell_mut((500, 500)).set_value_string("角落");
            umya_spreadsheet::writer::xlsx::write_writer(
                &wb,
                &mut std::fs::File::create(&doc.path).unwrap(),
            )
            .unwrap();
        }
        let before = doc.disk_bytes();
        let big = StructureOp::CopySheet {
            sheet: "大表".to_string(),
            name: None,
        };
        let err = doc.structure(&big, &[]).expect_err("超大表应拒绝复制");
        assert!(err.contains("表格过大"), "实际提示：{err}");
        assert_eq!(doc.disk_bytes(), before, "失败时原文件不得改变");
        assert!(!doc.store.has(doc.path.to_str().unwrap()), "失败不留影子");
        assert_eq!(dir_names(&doc.dir), vec!["huge.xlsx".to_string()]);
    }

    /* ----------------------------- 边界与失败 ----------------------------- */

    /// 只有 A1 有内容 / 只有 3 行数据的样本（用于「超出已用范围 = 空操作」的测试）
    fn make_sheet(path: &std::path::Path, sheet: &str, cells: &[(u32, u16, &str)]) {
        let mut wb = XlsxWriter::new();
        let target = wb.add_worksheet();
        target.set_name(sheet).unwrap();
        for (row, col, text) in cells {
            target.write_string(*row, *col, *text).unwrap();
        }
        wb.save(path).expect("生成测试 xlsx");
    }

    /// 插入行：`at` 超出已用范围 = 空操作（成功返回、内容不变、不留影子、不报 pending）
    #[test]
    fn insert_rows_outside_used_range_is_a_noop() {
        let doc = Doc::cells("insert-rows-outside", &[(0, 0, "唯一的格子")]);
        let disk = doc.disk_bytes();
        let before_rows = doc.rows("数据");

        let result = doc
            .structure(&insert_rows("数据", 100, 1), &[])
            .expect("超出已用范围的插入应成功（空操作）");
        assert!(!result.saved);
        assert!(!result.pending, "什么都没改，不该报 pending");
        assert_eq!(result.path, doc.path.to_string_lossy());
        assert!(!doc.store.has(doc.path.to_str().unwrap()), "不该留下影子");
        assert_eq!(doc.cell("数据", 0, 0), Data::String("唯一的格子".into()));
        assert_eq!(doc.cell("数据", 100, 0), Data::Empty, "不应真的插出内容");
        assert_eq!(doc.rows("数据"), before_rows, "工作表尺寸不应扩张");
        assert_eq!(doc.disk_bytes(), disk);

        // 空表上插入同样成功且什么也不发生
        let empty = Doc::empty("insert-rows-outside-empty", "empty.xlsx");
        make_sheet(&empty.path, "数据", &[]);
        empty
            .structure(&insert_rows("数据", 0, 3), &[])
            .expect("空表插入应成功");
        assert_eq!(empty.rows("数据"), 0);
        assert!(!empty.store.has(empty.path.to_str().unwrap()));
    }

    /// 插入列：`at` 超出已用范围 = 空操作
    #[test]
    fn insert_cols_outside_used_range_is_a_noop() {
        let doc = Doc::cells("insert-cols-outside", &[(0, 0, "唯一的格子")]);

        let result = doc
            .structure(&insert_cols("数据", 50, 1), &[])
            .expect("超出已用范围的插入列应成功（空操作）");
        assert!(!result.pending);
        assert_eq!(doc.cell("数据", 0, 0), Data::String("唯一的格子".into()));
        assert_eq!(doc.cell("数据", 0, 50), Data::Empty);
        assert_eq!(doc.cell("数据", 0, 1), Data::Empty, "不应真的插出新列");
    }

    /// 删除行：整段在已用范围之外 = 空操作
    #[test]
    fn delete_rows_outside_used_range_is_a_noop() {
        let doc = Doc::cells("delete-rows-outside", &[(0, 0, "唯一的格子")]);
        let before_rows = doc.rows("数据");

        let result = doc
            .structure(&delete_rows("数据", 100, 3), &[])
            .expect("整段在外的删除应成功（空操作）");
        assert!(!result.pending);
        assert_eq!(doc.cell("数据", 0, 0), Data::String("唯一的格子".into()));
        assert_eq!(doc.rows("数据"), before_rows, "不应删掉任何东西");
    }

    /// 删除行：部分重叠时「能删多少删多少」，不报错
    #[test]
    fn delete_rows_partially_overlapping_removes_only_existing_rows() {
        let doc = Doc::cells(
            "delete-rows-partial",
            &[(0, 0, "第一"), (1, 0, "第二"), (2, 0, "第三")],
        );
        assert_eq!(doc.rows("数据"), 3);

        // 从第 1 行开始删 5 行：只有第 1、2 行真实存在，删掉它们、第 0 行保留
        doc.structure(&delete_rows("数据", 1, 5), &[])
            .expect("部分重叠的删除应成功");
        assert_eq!(doc.cell("数据", 0, 0), Data::String("第一".into()));
        assert_eq!(doc.rows("数据"), 1, "只剩第 0 行");

        // 列的部分重叠同理
        let cols = Doc::cells("delete-cols-partial", &[(0, 0, "A"), (0, 1, "B"), (0, 2, "C")]);
        cols.structure(&delete_cols("数据", 1, 5), &[])
            .expect("部分重叠应成功");
        assert_eq!(cols.cell("数据", 0, 0), Data::String("A".into()));
        assert_eq!(cols.cell("数据", 0, 1), Data::Empty, "第 1 列之后都被删掉");
    }

    /// `count` 过大 / 非 xlsx / 表不存在：报错且原文件字节不变、无残留文件
    #[test]
    fn bad_requests_leave_file_untouched() {
        let doc = Doc::sample("bounds");
        let disk = doc.disk_bytes();

        // 每一条失败请求之后：原文件分毫不动、目录里只有样本文件
        let check_untouched = |err: String, want: &str| {
            assert!(err.contains(want), "错误信息应包含「{want}」：{err}");
            assert_eq!(doc.disk_bytes(), disk, "失败时原文件不得改变");
            assert_eq!(
                dir_names(&doc.dir),
                vec!["sample.xlsx".to_string()],
                "失败时不应留下备份/临时文件"
            );
            assert!(
                !doc.store.has(doc.path.to_str().unwrap()),
                "失败不该留下影子"
            );
        };

        let err = doc
            .structure(&delete_rows("数据", 0, MAX_OP_COUNT + 1), &[])
            .expect_err("count 过大应报错");
        check_untouched(err, "过大");

        let err = doc
            .structure(&insert_cols("数据", 0, MAX_OP_COUNT + 1), &[])
            .expect_err("count 过大应报错");
        check_untouched(err, "过大");

        let err = doc
            .structure(&delete_cols("不存在", 0, 1), &[])
            .expect_err("表不存在应报错");
        check_untouched(err, "不存在");

        let err = doc
            .structure(&insert_rows("不存在", 9999, 1), &[])
            .expect_err("表不存在应报错（越界不再是错误）");
        check_untouched(err, "不存在");

        // 非 xlsx：整簿重写会丢宏，必须拒绝
        let xlsm = doc.dir.join("sample.xlsm");
        std::fs::copy(&doc.path, &xlsm).unwrap();
        let err = office_shadow::apply_op(
            &doc.store,
            xlsm.to_str().unwrap(),
            &insert_rows("数据", 0, 1),
            &[],
        )
        .expect_err("xlsm 不应允许结构操作");
        assert!(err.contains("只有 .xlsx"), "实际提示：{err}");
        assert!(!doc.dir.join("sample.xlsm.bak").exists(), "拒绝时不应留下备份");

        // count = 0 视为 1：合法请求应真的插入 1 行（影子里的 A1 下移到 A2）
        let result = doc
            .structure(&insert_rows("数据", 0, 0), &[])
            .expect("count=0 应视为 1");
        assert!(result.pending);
        assert_eq!(doc.cell("数据", 1, 0), Data::String("名称".into()));
        assert_eq!(doc.cell("数据", 0, 0), Data::Empty, "插入的新行应为空");
        assert_eq!(doc.disk_cell("数据", 0, 0), Data::String("名称".into()));
    }

    /* --------------------------- 影子生命周期 --------------------------- */

    /// 结构操作后：磁盘字节完全不变、state 报 pending、丢弃后读回磁盘内容
    #[test]
    fn structure_keeps_disk_untouched_and_discard_restores_it() {
        let doc = Doc::sample("shadow-lifecycle");
        let disk = doc.disk_bytes();

        let state = doc.state().unwrap();
        assert!(!state.pending, "一开始没有未落盘改动");
        assert_eq!(state.sheets.len(), 2);

        doc.structure(&insert_rows("数据", 1, 1), &[])
            .expect("插入行应成功");

        // 1) 磁盘字节完全不变（逐字节比较）
        assert_eq!(doc.disk_bytes(), disk, "结构操作后磁盘字节必须完全不变");
        // 2) state 报 pending，且表列表来自影子
        let state = doc.state().unwrap();
        assert!(state.pending);
        assert_eq!(state.path, doc.path.to_string_lossy());
        assert_eq!(state.sheets.len(), 2);
        // 磁盘上的内容仍是原样
        assert_eq!(doc.disk_cell("数据", 1, 0), Data::String("苹果".into()));

        // 3) 丢弃后读到的又是磁盘内容
        assert!(doc.discard(), "确实丢掉了东西");
        assert!(!doc.state().unwrap().pending);
        assert_eq!(doc.cell("数据", 1, 0), Data::String("苹果".into()));
        assert_eq!(doc.rows("数据"), 5);
        assert!(!doc.discard(), "再丢一次就是 false");
    }

    /// 读路径（含 office.rs 的 info/rows 与 find/stats）都反映影子内容。
    ///
    /// 这一条**必须**用全局影子仓库：`office.rs` 的读路径只看全局那一份。
    #[test]
    fn reads_reflect_shadow_including_office_paths() {
        let doc = Doc::sample("reads-shadow");
        let p = doc.path.to_str().unwrap().to_string();

        // 先按磁盘内容读一次，把解析结果放进 SheetCache
        let cache = crate::commands::office::SheetCache::default();
        let before = crate::commands::office::rows_impl(&cache, &p, "数据", 0, 20)
            .expect("应能读窗口");
        assert_eq!(before.cells[1][0].v, "苹果");
        assert!(!crate::commands::office::info_impl(&p).unwrap().pending);

        office_shadow::apply_op(
            &office_shadow::SHADOWS,
            &p,
            &insert_rows("数据", 1, 1),
            &[],
        )
        .expect("插入行应成功");

        // 1) spreadsheet_info：pending = true，表列表来自影子
        let info = crate::commands::office::info_impl(&p).expect("info 应成功");
        assert!(info.pending, "info 应报有未落盘改动");
        assert!(info.editable);
        assert_eq!(info.sheets.len(), 2);

        // 2) spreadsheet_rows：**同一个 cache** 也必须看到影子内容
        //    （结构操作不改磁盘 mtime，靠缓存 key 里的影子版本号失效）
        let win = crate::commands::office::rows_impl(&cache, &p, "数据", 0, 20)
            .expect("应能读窗口");
        assert_eq!(win.cells[0][0].v, "名称");
        assert_eq!(win.cells[1][0].v, "", "插入的空行");
        assert_eq!(win.cells[2][0].v, "苹果", "原第 2 行下移");
        assert!(win.rows >= 6, "总行数应 +1");

        // 3) spreadsheet_find：命中新位置
        let hits = find_impl(&p, "苹果", None).expect("查找应成功");
        assert_eq!(hits.len(), 1);
        assert_eq!((hits[0].row, hits[0].col), (2, 0), "命中行号应是影子里的行号");

        // 4) spreadsheet_stats：统计的是影子内容（B 列 12.5 / 3 / 公式缓存 0 / 7）
        let stats = stats_impl(&p, "数据", &range(0, 1, 6, 1)).expect("统计应成功");
        assert_eq!(stats.numeric, 4);
        assert_eq!(stats.sum, 22.5);

        // 5) 新建的工作表同样立刻可见
        office_shadow::apply_op(
            &office_shadow::SHADOWS,
            &p,
            &StructureOp::AddSheet {
                name: Some("新表".to_string()),
            },
            &[],
        )
        .expect("新建表应成功");
        let info = crate::commands::office::info_impl(&p).unwrap();
        assert_eq!(
            info.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
            vec!["数据", "第二表", "新表"]
        );
        let hits = find_impl(&p, "苹果", Some(FindOptions { sheet: Some("数据".into()), ..Default::default() })).unwrap();
        assert_eq!(hits.len(), 1);

        // 收尾：丢弃影子，读到的又是磁盘内容（表列表回到 2 张、命中回到原位置）
        assert!(office_shadow::discard(&p));
        let info = crate::commands::office::info_impl(&p).unwrap();
        assert!(!info.pending);
        assert_eq!(info.sheets.len(), 2);
        let hits = find_impl(&p, "苹果", None).unwrap();
        assert_eq!(hits[0].row, 1, "丢弃后回到磁盘内容");
        // 同一个 cache 也要回到磁盘内容（版本号变回 0）
        let win = crate::commands::office::rows_impl(&cache, &p, "数据", 0, 20).unwrap();
        assert_eq!(win.cells[1][0].v, "苹果", "丢弃后缓存也必须回到磁盘内容");
        assert_eq!(win.cells[0][0].v, "名称");
    }

    /// 保存：磁盘 = 影子内容、pending 归 false、`.bak`（显式要求时）= 这批改动之前的磁盘字节；
    /// **保存后影子与撤销栈保留**，仍可一路撤销。
    #[test]
    fn save_lands_shadow_content_with_backup() {
        let doc = Doc::sample("save-shadow");
        let original = doc.disk_bytes();

        doc.structure(&insert_rows("数据", 1, 1), &[])
            .expect("插入行应成功");
        doc.structure(
            &StructureOp::AddSheet {
                name: Some("新表".to_string()),
            },
            &[],
        )
        .expect("新建表应成功");
        assert!(doc.state().unwrap().pending);

        let saved = doc.save(&[], true).expect("保存应成功");
        assert_eq!(saved.saved_cells, 0);
        assert_eq!(saved.path, doc.path.to_string_lossy());
        assert!(saved.size > 0 && saved.modified_at > 0);

        // 磁盘内容 = 之前影子的内容
        assert_eq!(doc.disk_cell("数据", 2, 0), Data::String("苹果".into()));
        assert_eq!(
            doc.disk_sheets(),
            vec!["数据", "第二表", "新表"],
            "新建的工作表已落盘"
        );
        // `.bak` = 操作前的原始字节
        assert_eq!(
            saved.backup_path.as_deref(),
            Some(doc.backup_path().to_string_lossy().as_ref())
        );
        assert_eq!(std::fs::read(doc.backup_path()).unwrap(), original);
        // pending 归 false；影子保留（撤销栈在里面），可撤销标志仍在
        let state = doc.state().unwrap();
        assert!(!state.pending, "保存后不是未落盘状态");
        assert!(state.can_undo, "保存不清空撤销栈");
        assert!(!state.can_redo);
        // 没有残留临时文件
        assert!(
            !dir_names(&doc.dir)
                .iter()
                .any(|n| n.starts_with(".MasterEdit-tmp-")),
            "不应残留临时文件：{:?}",
            dir_names(&doc.dir)
        );
    }

    /// 保存时 backup=false（也就是默认）：不生成 `.bak`
    #[test]
    fn save_backup_can_be_disabled() {
        let doc = Doc::sample("save-no-backup");
        doc.structure(&insert_rows("数据", 0, 1), &[])
            .expect("插入行应成功");

        let saved = doc.save(&[], false).expect("保存应成功");
        assert!(saved.backup_path.is_none());
        assert!(!doc.backup_path().exists(), "backup=false 时不应生成备份");
        assert_eq!(doc.disk_cell("数据", 1, 0), Data::String("名称".into()));
    }

    /* ------------------------------ 撤销/重做 ------------------------------ */

    /// 删除行 → 撤销：被删行的值回来了、行数恢复；can_undo / can_redo 正确
    #[test]
    fn undo_restores_deleted_rows() {
        let doc = Doc::sample("undo-delete");
        let before_rows = doc.rows("数据");

        doc.structure(&delete_rows("数据", 1, 1), &[])
            .expect("删除行应成功");
        assert_eq!(doc.rows("数据"), before_rows - 1);
        assert_eq!(doc.cell("数据", 1, 0), Data::String("香蕉".into()));
        let state = doc.state().unwrap();
        assert!(state.can_undo && !state.can_redo && state.pending);

        let result = office_shadow::undo(&doc.store, doc.path.to_str().unwrap())
            .expect("应能撤销");
        assert_eq!(result.path, doc.path.to_string_lossy());
        assert!(!result.can_undo && result.can_redo);
        assert!(!result.pending, "退回初始状态 → 与磁盘一致");
        assert_eq!(
            doc.rows("数据"),
            before_rows,
            "行数应恢复到删除前"
        );
        assert_eq!(
            doc.cell("数据", 1, 0),
            Data::String("苹果".into()),
            "被删行的值应该回来"
        );

        // 重做：又回到删除后的状态
        let result = office_shadow::redo(&doc.store, doc.path.to_str().unwrap()).expect("应能重做");
        assert!(result.can_undo && !result.can_redo && result.pending);
        assert_eq!(doc.rows("数据"), before_rows - 1);
        assert_eq!(doc.cell("数据", 1, 0), Data::String("香蕉".into()));
    }

    /// 快照时机：`edits` 之后、结构 op 之前 —— 撤销结构操作不会连带回退单元格编辑
    #[test]
    fn snapshot_after_edits_keeps_cell_edits_on_undo() {
        let doc = Doc::sample("undo-timing");
        let before_rows = doc.rows("数据");
        let edits = [edit("数据", 0, 0, "text", "x")];

        doc.structure(&insert_rows("数据", 1, 1), &edits)
            .expect("插入行应成功");
        assert_eq!(doc.cell("数据", 0, 0), Data::String("x".into()));
        assert_eq!(doc.cell("数据", 2, 0), Data::String("苹果".into()));
        assert_eq!(doc.rows("数据"), before_rows + 1);

        let result = office_shadow::undo(&doc.store, doc.path.to_str().unwrap())
            .expect("应能撤销结构操作");
        assert!(result.pending, "编辑仍在影子里 → 与磁盘不同");
        assert_eq!(
            doc.cell("数据", 0, 0),
            Data::String("x".into()),
            "撤销结构操作不能把单元格编辑一起回退"
        );
        assert_eq!(doc.rows("数据"), before_rows, "行数回到插入前");
        assert_eq!(doc.cell("数据", 1, 0), Data::String("苹果".into()));
    }

    /// 保存后仍可撤销：影子回到操作前、pending 重新为 true、磁盘保持刚保存的内容
    #[test]
    fn undo_after_save_keeps_disk_and_reverts_shadow() {
        let doc = Doc::sample("undo-after-save");
        let original = doc.disk_bytes();

        doc.structure(&insert_rows("数据", 1, 1), &[])
            .expect("插入行应成功");
        doc.save(&[], false).expect("保存应成功");
        let saved_disk = doc.disk_bytes();
        assert_ne!(saved_disk, original, "保存确实写了盘");
        assert!(!doc.state().unwrap().pending);
        assert_eq!(doc.disk_cell("数据", 2, 0), Data::String("苹果".into()));

        let result =
            office_shadow::undo(&doc.store, doc.path.to_str().unwrap()).expect("保存后仍应能撤销");
        assert!(result.pending, "撤销后重新变脏");
        assert!(!result.can_undo && result.can_redo);
        // 影子回到操作前
        assert_eq!(doc.cell("数据", 1, 0), Data::String("苹果".into()));
        assert_eq!(doc.disk_cell("数据", 2, 0), Data::String("苹果".into()));
        // 磁盘仍是刚保存的内容（撤销不写盘）
        assert_eq!(doc.disk_bytes(), saved_disk, "撤销不该改磁盘");

        // 再按 Ctrl+S：把撤销后的状态写回去
        doc.save(&[], false).expect("再次保存应成功");
        assert!(!doc.state().unwrap().pending);
        assert_eq!(doc.disk_cell("数据", 1, 0), Data::String("苹果".into()));
        let after = doc.disk_bytes();
        let mut wb = open_workbook_auto_from_rs(Cursor::new(after)).unwrap();
        assert_eq!(
            rows_of_wb(&mut wb, "数据"),
            rows_of_wb(&mut open_workbook_auto_from_rs(Cursor::new(original)).unwrap(), "数据"),
            "两次保存后行数与最初一致"
        );
    }

    /// 空操作不压栈：越界插入后没有可撤销的步骤（连影子都不建）
    #[test]
    fn noop_structure_does_not_push_undo() {
        let doc = Doc::cells("undo-noop", &[(0, 0, "唯一的格子")]);

        doc.structure(&insert_rows("数据", 100, 1), &[])
            .expect("越界插入应成功");
        let state = doc.state().unwrap();
        assert!(!state.can_undo, "空操作不该产生可撤销的步骤");
        assert!(!state.can_redo);
        assert!(!state.pending);
        assert_eq!(
            office_shadow::undo(&doc.store, doc.path.to_str().unwrap())
                .expect_err("没有可撤销的改动"),
            "没有可撤销的改动"
        );

        // 带上 edits 时内容确实变了（pending 为 true），但结构本身没改 → 仍然没有结构步；
        // 那些编辑属于前端自己的「cells」撤销步，后端替它压栈会让两条栈错位。
        let with_edits = Doc::cells("undo-noop-edits", &[(0, 0, "唯一的格子")]);
        with_edits
            .structure(&insert_rows("数据", 100, 1), &[edit("数据", 0, 0, "text", "改过")])
            .expect("越界插入 + 编辑应成功");
        assert!(with_edits.state().unwrap().pending, "编辑改了内容");
        assert!(
            !with_edits.state().unwrap().can_undo,
            "结构没改 → 不压结构快照"
        );
        assert_eq!(with_edits.cell("数据", 0, 0), Data::String("改过".into()));
    }

    /// 企业加密文件：撤销/重做只改内存，保存后仍加密、解密可读
    #[test]
    fn encrypted_undo_redo_and_save_keep_encryption() {
        let doc = Doc::cells("encrypted-undo", &[(0, 0, "第一"), (1, 0, "第二")]);
        let plain = doc.disk_bytes();
        let wrapped = wrap_encrypted(&plain);
        std::fs::write(&doc.path, &wrapped).unwrap();

        doc.structure(&insert_rows("数据", 1, 1), &[])
            .expect("加密文件应能改结构");
        assert_eq!(doc.disk_bytes(), wrapped, "结构操作后磁盘仍是原密文");
        assert!(doc.store.is_pending(doc.path.to_str().unwrap()));

        // 撤销 / 重做都只改内存
        office_shadow::undo(&doc.store, doc.path.to_str().unwrap()).expect("应能撤销");
        assert_eq!(doc.disk_bytes(), wrapped, "撤销不碰磁盘");
        assert_eq!(doc.cell("数据", 1, 0), Data::String("第二".into()), "回到插入前");
        office_shadow::redo(&doc.store, doc.path.to_str().unwrap()).expect("应能重做");
        assert_eq!(doc.disk_bytes(), wrapped);
        assert_eq!(doc.cell("数据", 2, 0), Data::String("第二".into()));

        // 保存：仍是加密格式、解密可读、撤销栈还在
        let saved = doc.save(&[], true).expect("保存应成功");
        assert_eq!(saved.saved_cells, 0);
        let out = doc.disk_bytes();
        assert_eq!(&out[..4], &[0xE0, 0xA8, 0x91, 0xE7], "输出必须仍是加密格式");
        assert!(
            !out.windows(4).any(|w| w == b"PK\x03\x04"),
            "加密文件里不得出现明文 zip 内容"
        );
        assert_eq!(
            std::fs::read(doc.backup_path()).unwrap(),
            wrapped,
            "备份应是原文件的密文字节"
        );
        assert!(doc.state().unwrap().can_undo, "保存后撤销栈仍在");
        let back = esafenet::decrypt_esafenet(&out).expect("落盘内容应能解密");
        let mut wb = open_workbook_auto_from_rs(Cursor::new(back)).unwrap();
        assert_eq!(cell_of(&mut wb, "数据", 2, 0), Data::String("第二".into()));
        assert_eq!(cell_of(&mut wb, "数据", 0, 0), Data::String("第一".into()));
    }

    /// 连续多次结构操作累积在同一个影子里，最后一次保存全部落盘
    #[test]
    fn multiple_operations_accumulate_until_save() {
        let doc = Doc::sample("accumulate");
        let disk = doc.disk_bytes();

        // 插入行 → 新建表 → 删除列
        doc.structure(&insert_rows("数据", 1, 1), &[])
            .expect("插入行应成功");
        doc.structure(
            &StructureOp::AddSheet {
                name: Some("新表".to_string()),
            },
            &[],
        )
        .expect("新建表应成功");
        doc.structure(&delete_cols("数据", 3, 1), &[])
            .expect("删除列应成功");

        assert_eq!(doc.sheets(), vec!["数据", "第二表", "新表"]);
        assert_eq!(doc.cell("数据", 2, 0), Data::String("苹果".into()));
        assert_eq!(doc.cell("数据", 4, 3), Data::Empty, "D 列的 x 已被删");
        assert_eq!(doc.disk_bytes(), disk, "三次操作都没碰磁盘");
        assert_eq!(doc.disk_sheets(), vec!["数据", "第二表"]);

        // 一次保存全部落盘
        doc.save(&[], true).expect("保存应成功");
        assert_eq!(doc.disk_sheets(), vec!["数据", "第二表", "新表"]);
        assert_eq!(doc.disk_cell("数据", 2, 0), Data::String("苹果".into()));
        assert_eq!(doc.disk_cell("数据", 4, 3), Data::Empty);
        assert_eq!(std::fs::read(doc.backup_path()).unwrap(), disk);
        assert!(!doc.state().unwrap().pending);
    }

    /// 企业加密文件：结构操作后磁盘仍是原密文；保存后仍加密、解密可读、`.bak` = 原密文
    #[test]
    fn encrypted_shadow_save_keeps_encryption() {
        let doc = Doc::cells("encrypted-shadow", &[(0, 0, "第一"), (1, 0, "第二")]);
        let plain = doc.disk_bytes();
        let wrapped = wrap_encrypted(&plain);
        std::fs::write(&doc.path, &wrapped).unwrap();

        doc.structure(&insert_rows("数据", 1, 1), &[])
            .expect("加密工作簿应能改结构");
        assert!(
            doc.store.is_pending(doc.path.to_str().unwrap()),
            "加密文件同样只在影子里改"
        );
        assert_eq!(doc.disk_bytes(), wrapped, "结构操作后磁盘仍是原密文");

        let saved = doc.save(&[], true).expect("保存应成功");
        assert_eq!(saved.saved_cells, 0);

        let out = doc.disk_bytes();
        assert_eq!(&out[..4], &[0xE0, 0xA8, 0x91, 0xE7], "输出必须仍是加密格式");
        assert!(
            !out.windows(4).any(|w| w == b"PK\x03\x04"),
            "加密文件里不得出现明文 zip 内容"
        );
        assert_eq!(
            std::fs::read(doc.backup_path()).unwrap(),
            wrapped,
            "备份应是原文件的密文字节"
        );

        // 解密后能读到新结构（原第 1 行下移到第 2 行）
        let back = esafenet::decrypt_esafenet(&out).expect("落盘内容应能解密");
        let mut wb = open_workbook_auto_from_rs(Cursor::new(back)).unwrap();
        assert_eq!(
            cell_of(&mut wb, "数据", 2, 0),
            Data::String("第二".into()),
            "插入行后原第 2 行下移到第 3 行"
        );
        assert_eq!(cell_of(&mut wb, "数据", 0, 0), Data::String("第一".into()));
    }

    /* ------------------------------ 选区统计 ------------------------------ */

    fn range(r0: u32, c0: u32, r1: u32, c1: u32) -> CellRange {
        CellRange {
            start_row: r0,
            start_col: c0,
            end_row: r1,
            end_col: c1,
        }
    }

    /// 混合文本 / 数字 / 数字文本 / 空格的矩形
    #[test]
    fn stats_on_mixed_rectangle() {
        let dir = test_dir("stats-mixed");
        let path = dir.join("stats.xlsx");
        {
            let mut wb = XlsxWriter::new();
            let sheet = wb.add_worksheet();
            sheet.set_name("数据").unwrap();
            sheet.write_number(0, 0, 1.0).unwrap(); // A1：数值
            sheet.write_string(0, 1, "2.5").unwrap(); // B1：数字文本，计入数值
            sheet.write_string(0, 2, "abc").unwrap(); // C1：文本
            sheet.write_boolean(1, 0, true).unwrap(); // A2：布尔，非空但不是数值
            // B2 留空
            sheet.write_string(1, 2, "文本").unwrap(); // C2
            sheet.write_number(1, 3, 10.0).unwrap(); // D2 在选区之外
            wb.save(&path).unwrap();
        }
        let p = path.to_str().unwrap();

        let stats = stats_impl(p, "数据", &range(0, 0, 1, 2)).unwrap();
        assert_eq!(stats.cells, 6, "2 行 × 3 列");
        assert_eq!(
            stats.non_empty, 5,
            "A1 B1 C1 A2 C2 非空（布尔也算非空），B2 是空格"
        );
        assert_eq!(stats.numeric, 2, "1 与 \"2.5\"");
        assert_eq!(stats.sum, 3.5);
        assert_eq!(stats.average, Some(1.75));
        assert_eq!(stats.min, Some(1.0));
        assert_eq!(stats.max, Some(2.5));

        // 单格选区：文本格
        let single = stats_impl(p, "数据", &range(0, 2, 0, 2)).unwrap();
        assert_eq!(single.cells, 1);
        assert_eq!(single.non_empty, 1);
        assert_eq!(single.numeric, 0);
        assert_eq!(single.average, None);

        // 只框住空白的格子：cells 有、其余全零
        let blank = stats_impl(p, "数据", &range(1, 1, 1, 1)).unwrap();
        assert_eq!(blank.cells, 1);
        assert_eq!(blank.non_empty, 0);
        assert_eq!(blank.numeric, 0);
        assert_eq!(blank.sum, 0.0);
    }

    /// 全文本区域：numeric = 0、sum = 0、average/min/max 为 null
    #[test]
    fn stats_all_text_has_no_numbers() {
        let dir = test_dir("stats-text");
        let path = dir.join("text.xlsx");
        {
            let mut wb = XlsxWriter::new();
            let sheet = wb.add_worksheet();
            sheet.set_name("数据").unwrap();
            sheet.write_string(0, 0, "甲").unwrap();
            sheet.write_string(0, 1, "十二").unwrap();
            sheet.write_string(1, 0, "-").unwrap();
            wb.save(&path).unwrap();
        }
        let stats = stats_impl(path.to_str().unwrap(), "数据", &range(0, 0, 1, 1)).unwrap();
        assert_eq!(stats.cells, 4);
        assert_eq!(stats.non_empty, 3);
        assert_eq!(stats.numeric, 0);
        assert_eq!(stats.sum, 0.0);
        assert_eq!(stats.average, None);
        assert_eq!(stats.min, None);
        assert_eq!(stats.max, None);
    }

    /// 区域超出工作表范围：超出部分忽略，不报错；起止颠倒自动交换
    #[test]
    fn stats_clamps_to_sheet_and_swaps_reversed_range() {
        let dir = test_dir("stats-clamp");
        let path = dir.join("clamp.xlsx");
        {
            let mut wb = XlsxWriter::new();
            let sheet = wb.add_worksheet();
            sheet.set_name("数据").unwrap();
            sheet.write_number(0, 0, 2.0).unwrap();
            sheet.write_number(1, 0, 3.0).unwrap();
            wb.save(&path).unwrap();
        }
        let p = path.to_str().unwrap();

        // 选区远超已用范围（2 行 × 1 列）
        let stats = stats_impl(p, "数据", &range(0, 0, 999, 5)).unwrap();
        assert_eq!(stats.non_empty, 2, "超出范围的部分忽略");
        assert_eq!(stats.numeric, 2);
        assert_eq!(stats.sum, 5.0);
        assert_eq!(stats.cells, 2, "cells 只统计落在工作表已用范围内的格子");

        // 起止颠倒 → 与正向一致
        let reversed = stats_impl(p, "数据", &range(1, 0, 0, 0)).unwrap();
        let normal = stats_impl(p, "数据", &range(0, 0, 1, 0)).unwrap();
        assert_eq!(reversed, normal, "颠倒的选区应与正向结果完全一致");
        assert_eq!(reversed.sum, 5.0);

        // 完全在已用范围之外
        let outside = stats_impl(p, "数据", &range(100, 100, 101, 101)).unwrap();
        assert_eq!(outside, RangeStats::default());

        // 表不存在
        let err = stats_impl(p, "缺失表", &range(0, 0, 1, 1)).expect_err("表不存在应报错");
        assert!(err.contains("缺失表"), "实际提示：{err}");
    }

    /// 过大选区报错（含颠倒的选区：先交换再判上限）
    #[test]
    fn stats_rejects_oversized_range() {
        let dir = test_dir("stats-big");
        let path = dir.join("big.xlsx");
        make_sample(&path);
        let p = path.to_str().unwrap();

        let err = stats_impl(p, "数据", &range(0, 0, 1_100, 1_000)).expect_err("过大选区应报错");
        assert!(err.contains("选区过大"), "实际提示：{err}");

        let err = stats_impl(p, "数据", &range(1_100, 1_000, 0, 0)).expect_err("过大选区应报错");
        assert!(err.contains("选区过大"), "实际提示：{err}");

        // 恰好在上限内：不报错
        let stats = stats_impl(p, "数据", &range(0, 0, 999, 999)).unwrap();
        assert!(stats.cells <= MAX_STATS_CELLS as usize);
    }

    /// 真实样本（本机存在时才跑）：企业加密 xlsx 上的查找与统计
    #[test]
    fn real_encrypted_find_and_stats_if_present() {
        let real = r"Z:\D\mywork\01_project\007_topology_identification\06_data\voltage_daily_v3\2026-03-01.xlsx";
        if !Path::new(real).exists() {
            return;
        }
        let hits = find_impl(real, "A", None).expect("真实加密工作簿应能查找");
        println!("真实样本查找命中 {} 条，首条 {:?}", hits.len(), hits.first());
        let sheet = {
            let (bytes, encrypted) = read_workbook_bytes(Path::new(real)).unwrap();
            assert!(encrypted, "真实样本应是企业加密文档");
            let wb = parse_workbook(&bytes).unwrap();
            wb.sheet_names().first().cloned().unwrap()
        };
        let stats = stats_impl(real, &sheet, &range(0, 0, 200, 50)).expect("真实加密工作簿应能统计");
        println!("真实样本首表统计：{stats:?}");
        assert!(stats.cells > 0);
    }
}
