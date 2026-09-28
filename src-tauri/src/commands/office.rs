//! 办公文档（电子表格）只读解析：xlsx / xls / xlsb / ods。
//!
//! 设计要点：
//! - 解析交给 calamine（纯 Rust，无外部进程、无 JS 依赖），只在打开表格时按需加载；
//! - 数据按「行窗口」返回，配合前端虚拟滚动，避免大表一次性序列化成几十 MB JSON；
//! - 解析结果按 (路径, 修改时间, 工作表) 做小容量缓存，滚动时无需重复解析；
//! - 本模块**只读**，绝不写回原文件（编辑能力属于后续阶段）。
//!
//! 限制与已知取舍：
//! - 1904 日期系统（老式 Mac 工作簿）无法从 calamine 读取标志位，日期会整体偏移 1462 天；
//! - 公式只返回文本（含前导 `=`），不做计算，数值仍是工作簿里缓存的显示值；
//! - 超过 `MAX_COLS` 的列、超过上限的单元格文本会被截断，并通过 `truncated` 告知前端。

use crate::commands::esafenet;
use crate::commands::file::modified_ms;
use crate::commands::image::civil_from_days;
use calamine::{open_workbook_auto_from_rs, CellErrorType, Data, Range, Reader, Sheets};
use serde::Serialize;
use std::io::Cursor;
use std::path::Path;
use std::sync::Mutex;

/// 单次请求返回的最大行数
const MAX_WINDOW_ROWS: usize = 1000;
/// 单次请求返回的最大单元格数（行数会据此进一步收窄）
const MAX_WINDOW_CELLS: usize = 20_000;
/// 单个工作表最多展示的列数（xlsx 理论上限 16384，查看场景无需那么多）
const MAX_COLS: usize = 1024;
/// 单元格文本上限，防止超长文本撑爆传输
const MAX_CELL_CHARS: usize = 2000;
/// 缓存条目上限（小工作表最多同时缓存这么多张）
const CACHE_ENTRIES: usize = 3;
/// 缓存的总单元格预算：超过后淘汰最久未用的条目。
/// 注意「单张超大工作表」会被保留（只在条目数 > 1 时淘汰），
/// 否则百万行的工作表每个滚动窗口都要重新解析整表。
const CACHE_CELL_BUDGET: usize = 1_500_000;

/// 工作表元信息（行列为 0 表示尚未加载，前端在打开该表时才拿到真实尺寸）
#[derive(Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct SheetMeta {
    pub name: String,
}

/// `spreadsheet_info` 返回：文件信息 + 工作表列表
#[derive(Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetInfo {
    pub path: String,
    pub sheets: Vec<SheetMeta>,
    pub modified_at: u64,
    pub size: u64,
    /// 是否为「企业透明加密」文档（已在内存中解密，仅用于提示用户）
    pub encrypted: bool,
}

/// 单元格：`v` 为显示文本，`t` 为类型，`f` 为公式（仅含公式的单元格才有）
#[derive(Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct SheetCell {
    pub v: String,
    pub t: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub f: Option<String>,
}

/// `spreadsheet_rows` 返回：一个行窗口
#[derive(Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetWindow {
    pub sheet: String,
    /// 该表总行数（绝对行号 + 1，与 Excel 的行号一致）
    pub rows: usize,
    /// 该表展示列数（已按 MAX_COLS 截断）
    pub cols: usize,
    /// 本窗口起始行（0 起，绝对行号）
    pub start: usize,
    pub cells: Vec<Vec<SheetCell>>,
    /// 是否因列数/文本长度上限被截断
    pub truncated: bool,
}

/// 已解析的工作表缓存
struct CachedSheet {
    key: String,
    modified_ms: u64,
    values: Range<Data>,
    formulas: Option<Range<String>>,
    /// 该表的单元格数，用于总预算淘汰
    cells: usize,
}

/// 常驻状态：key = `小写路径|工作表名` -> 解析结果（按最近使用排序，末尾最新）
#[derive(Default)]
pub struct SheetCache(pub Mutex<SheetCacheInner>);

/// 缓存内部结构：`entries` 末尾为最近使用
#[derive(Default)]
pub struct SheetCacheInner {
    entries: Vec<CachedSheet>,
}

impl SheetCacheInner {
    /// 命中则取出并把条目移到末尾（标记为最近使用）
    fn take(&mut self, key: &str, modified_ms: u64) -> Option<(Range<Data>, Option<Range<String>>)> {
        let index = self
            .entries
            .iter()
            .position(|e| e.key == key && e.modified_ms == modified_ms)?;
        let entry = self.entries.remove(index);
        let out = (entry.values.clone(), entry.formulas.clone());
        self.entries.push(entry);
        Some(out)
    }

    /// 写入缓存并按「条目数 + 总单元格预算」淘汰最久未用的条目
    fn put(&mut self, entry: CachedSheet) {
        self.entries.retain(|e| e.key != entry.key);
        self.entries.push(entry);
        while self.entries.len() > CACHE_ENTRIES {
            self.entries.remove(0);
        }
        while self.entries.len() > 1
            && self.entries.iter().map(|e| e.cells).sum::<usize>() > CACHE_CELL_BUDGET
        {
            self.entries.remove(0);
        }
    }
}

/* ------------------------------------------------------------------ */
/* 单元格取值与格式化                                                  */
/* ------------------------------------------------------------------ */

/// 截断超长文本（按字符边界，避免多字节字符被切开）
fn truncate_text(s: &str) -> String {
    if s.chars().count() <= MAX_CELL_CHARS {
        return s.to_string();
    }
    let mut out: String = s.chars().take(MAX_CELL_CHARS).collect();
    out.push('…');
    out
}

/// 浮点显示：整数不带小数点，其余用最短往返表示
fn format_float(f: f64) -> String {
    if !f.is_finite() {
        return format!("{f}");
    }
    let text = if f.fract() == 0.0 && f.abs() < 1e15 {
        format!("{f:.0}")
    } else {
        format!("{f}")
    };
    // -0 归一化为 0
    if text == "-0" {
        "0".to_string()
    } else {
        text
    }
}

/// 一天内的小数部分 → HH:MM:SS（四舍五入到秒）
fn format_time_of_day(frac: f64) -> String {
    let total = (frac * 86_400.0).round() as i64;
    let total = total.rem_euclid(86_400);
    format!("{:02}:{:02}:{:02}", total / 3600, (total % 3600) / 60, total % 60)
}

/// 时长（Excel 的 [h]:mm:ss 语义，小时可超过 24）
fn format_duration(serial: f64) -> String {
    let total = (serial * 86_400.0).round() as i64;
    let sign = if total < 0 { "-" } else { "" };
    let t = total.abs();
    format!("{sign}{}:{:02}:{:02}", t / 3600, (t % 3600) / 60, t % 60)
}

/// Excel 序列日期 → 显示文本。
///
/// 1900 日期系统：序列值 61 起以 1899-12-30 为原点；1..59 需回退一天
/// （Excel 把 1900 当闰年，多出的 1900-02-29 占用了序列值 60）。
/// 已用 .NET 逐点核对：1→1900-01-01、59→1900-02-28、61→1900-03-01、25569→1970-01-01、44927→2023-01-01。
fn format_serial(serial: f64, is_duration: bool) -> String {
    if !serial.is_finite() {
        return format!("{serial}");
    }
    if is_duration {
        return format_duration(serial);
    }
    if serial < 0.0 {
        return format_float(serial);
    }
    // 纯时间（小于 1 天）
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

/// Excel 错误码 → 与 Excel 一致的显示文本
fn error_text(e: &CellErrorType) -> &'static str {
    match e {
        CellErrorType::Div0 => "#DIV/0!",
        CellErrorType::NA => "#N/A",
        CellErrorType::Name => "#NAME?",
        CellErrorType::Null => "#NULL!",
        CellErrorType::Num => "#NUM!",
        CellErrorType::Ref => "#REF!",
        CellErrorType::Value => "#VALUE!",
        CellErrorType::GettingData => "#GETTING_DATA",
    }
}

/// calamine 单元格 → (显示文本, 类型)
fn cell_text(data: &Data) -> (String, &'static str) {
    match data {
        Data::Empty => (String::new(), "empty"),
        Data::String(s) => (truncate_text(s), "text"),
        Data::Int(i) => (i.to_string(), "number"),
        Data::Float(f) => (format_float(*f), "number"),
        Data::Bool(b) => ((if *b { "true" } else { "false" }).to_string(), "bool"),
        Data::DateTime(dt) => (
            format_serial(dt.as_f64(), dt.is_duration()),
            if dt.is_duration() { "text" } else { "date" },
        ),
        Data::DateTimeIso(s) => (truncate_text(s), "date"),
        Data::DurationIso(s) => (truncate_text(s), "text"),
        Data::Error(e) => (error_text(e).to_string(), "error"),
    }
}

/* ------------------------------------------------------------------ */
/* 解析与缓存                                                          */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/* 文件读取：企业加密解密 + 格式识别                                    */
/* ------------------------------------------------------------------ */

/// 读取工作簿字节。企业透明加密（亿赛通等）文档在**内存中**解密后交给 calamine，
/// 不写临时文件，避免把明文副本落到磁盘上。
fn read_workbook_bytes(p: &Path) -> Result<(Vec<u8>, bool), String> {
    let raw = std::fs::read(p).map_err(|e| format!("读取文件失败: {e}"))?;
    match esafenet::decrypt_esafenet(&raw) {
        Some(plain) => Ok((plain, true)),
        None => Ok((raw, false)),
    }
}

/// 从内存解析工作簿：calamine 按内容自动识别 xls / xlsx / xlsb / ods
fn parse_workbook(bytes: &[u8]) -> Result<Sheets<Cursor<&[u8]>>, String> {
    open_workbook_auto_from_rs(Cursor::new(bytes)).map_err(|e| e.to_string())
}

/// 把底层解析错误翻译成用户能照着做的提示。
/// 常见情况：需要密码的 Office 文档 / 旧版二进制格式（都是 CFB 容器）、
/// 系统导出的「假 Excel」（真实内容是 HTML 表格）、被加密软件处理过或已损坏的文件。
fn describe_parse_failure(bytes: &[u8], err: &str) -> String {
    if bytes.len() >= 4 && bytes[..4] == [0xD0, 0xCF, 0x11, 0xE0] {
        return format!(
            "该文件是加密的 Office 文档（打开需要密码）或旧版二进制格式，暂时无法解析。\n\
             请先用 Excel / WPS 打开，另存为不加密的 .xlsx 后重试。\n（底层错误：{err}）"
        );
    }
    let head_len = bytes.len().min(2048);
    let head = String::from_utf8_lossy(&bytes[..head_len]).to_lowercase();
    if head.contains("<html") || head.contains("<table") || head.contains("<!doctype") {
        return format!(
            "该文件的真实内容是网页表格（不少系统导出的“Excel”其实是 HTML），扩展名被改成了表格格式。\n\
             请让导出方另存为真正的 .xlsx，或把文件改名为 .html 后用浏览器打开。\n（底层错误：{err}）"
        );
    }
    if bytes.len() >= 2 && bytes[..2] != *b"PK" {
        return format!(
            "文件头不是 Excel 标准格式：可能被加密软件处理过、下载不完整或已损坏。\n\
             可尝试用 Excel / WPS 打开后另存，或右键标签选择「用系统默认程序打开」。\n（底层错误：{err}）"
        );
    }
    format!("解析表格失败：{err}")
}

/// 取（或解析并缓存）某个工作表的值与公式
fn open_sheet(
    path: &str,
    sheet: &str,
    cache: &SheetCache,
) -> Result<(Range<Data>, Option<Range<String>>, u64), String> {
    let p = Path::new(path);
    let meta = std::fs::metadata(p).map_err(|e| format!("读取文件信息失败: {e}"))?;
    if !meta.is_file() {
        return Err("目标不是有效的文件".to_string());
    }
    let mtime = modified_ms(&meta);
    let key = format!("{}|{}", path.to_lowercase(), sheet);

    // 命中缓存（且文件未被外部修改）时直接复用
    if let Ok(mut inner) = cache.0.lock() {
        if let Some(hit) = inner.take(&key, mtime) {
            return Ok((hit.0, hit.1, mtime));
        }
    }

    let (bytes, _) = read_workbook_bytes(p)?;
    let mut workbook = parse_workbook(&bytes).map_err(|e| describe_parse_failure(&bytes, &e))?;
    let values = workbook
        .worksheet_range(sheet)
        .map_err(|e| format!("读取工作表「{sheet}」失败: {e}"))?;
    // 公式为可选能力：xls / ods 等格式可能不支持，失败即视为无公式
    let formulas = workbook.worksheet_formula(sheet).ok();

    // 解析结果一律进缓存（大表也缓存，否则每个滚动窗口都要重新解析整表）；
    // 内存由「条目数 + 总单元格预算」控制
    let (height, width) = values.get_size();
    if let Ok(mut inner) = cache.0.lock() {
        inner.put(CachedSheet {
            key,
            modified_ms: mtime,
            values: values.clone(),
            formulas: formulas.clone(),
            cells: width.saturating_mul(height),
        });
    }
    Ok((values, formulas, mtime))
}

fn info_impl(path: &str) -> Result<SpreadsheetInfo, String> {
    let p = Path::new(path);
    let meta = std::fs::metadata(p).map_err(|e| format!("读取文件信息失败: {e}"))?;
    if !meta.is_file() {
        return Err("目标不是有效的文件".to_string());
    }
    let (bytes, encrypted) = read_workbook_bytes(p)?;
    let workbook = parse_workbook(&bytes).map_err(|e| describe_parse_failure(&bytes, &e))?;
    let sheets = workbook
        .sheet_names()
        .into_iter()
        .map(|name| SheetMeta { name })
        .collect();
    Ok(SpreadsheetInfo {
        path: path.to_string(),
        sheets,
        modified_at: modified_ms(&meta),
        size: meta.len(),
        encrypted,
    })
}

fn rows_impl(
    cache: &SheetCache,
    path: &str,
    sheet: &str,
    start: usize,
    count: usize,
) -> Result<SpreadsheetWindow, String> {
    let (values, formulas, _) = open_sheet(path, sheet, cache)?;
    // 注意：calamine 的 get_size() 返回 (高, 宽)；Range 从首个非空单元格开始，
    // 因此统一用绝对坐标 get_value() 取值，行号/列标才能与 Excel 对齐
    let (height, width) = values.get_size();
    let (end_row, end_col) = values
        .end()
        .map(|(r, c)| (r as usize, c as usize))
        .unwrap_or((0, 0));
    let total_rows = if height == 0 { 0 } else { end_row + 1 };
    let total_cols = if width == 0 { 0 } else { end_col + 1 };
    let cols = total_cols.min(MAX_COLS);
    let truncated_cols = total_cols > MAX_COLS;

    // 单次窗口同时受行数与单元格数限制，避免一次返回过多数据
    let mut take = count.clamp(1, MAX_WINDOW_ROWS);
    if cols > 0 {
        take = take.min((MAX_WINDOW_CELLS / cols).max(1));
    }
    let start = start.min(total_rows);
    let end = (start + take).min(total_rows);

    let mut cells = Vec::with_capacity(end.saturating_sub(start));
    let mut truncated_text = false;
    for r in start..end {
        let mut row = Vec::with_capacity(cols);
        for c in 0..cols {
            let position = (r as u32, c as u32);
            let (v, t) = match values.get_value(position) {
                Some(data) => cell_text(data),
                None => (String::new(), "empty"),
            };
            if v.chars().count() > MAX_CELL_CHARS {
                truncated_text = true;
            }
            // 公式区域与值区域的起点可能不同，同样按绝对坐标取值
            let f = formulas
                .as_ref()
                .and_then(|range| range.get_value(position))
                .map(|s| s.trim())
                .filter(|s| !s.is_empty())
                .map(|s| {
                    if s.starts_with('=') {
                        truncate_text(s)
                    } else {
                        format!("={}", truncate_text(s))
                    }
                });
            row.push(SheetCell { v, t, f });
        }
        cells.push(row);
    }

    Ok(SpreadsheetWindow {
        sheet: sheet.to_string(),
        rows: total_rows,
        cols,
        start,
        cells,
        truncated: truncated_cols || truncated_text,
    })
}

/* ------------------------------------------------------------------ */
/* Tauri 命令                                                          */
/* ------------------------------------------------------------------ */

/// 读取表格文件的工作表列表（不做整表解析，打开很快）
#[tauri::command]
pub async fn spreadsheet_info(path: String) -> Result<SpreadsheetInfo, String> {
    tauri::async_runtime::spawn_blocking(move || info_impl(&path))
        .await
        .map_err(|e| format!("解析表格失败: {e}"))?
}

/// 按行窗口读取工作表内容（配合前端虚拟滚动）
#[tauri::command]
pub async fn spreadsheet_rows(
    app: tauri::AppHandle,
    path: String,
    sheet: String,
    start: usize,
    count: usize,
) -> Result<SpreadsheetWindow, String> {
    tauri::async_runtime::spawn_blocking(move || {
        use tauri::Manager;
        let cache = app.state::<SheetCache>();
        rows_impl(&cache, &path, &sheet, start, count)
    })
    .await
    .map_err(|e| format!("解析表格失败: {e}"))?
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;

    /// 生成一个测试用 xlsx：两个工作表、中文表名、数字/布尔/公式/长文本
    fn make_workbook(name: &str) -> std::path::PathBuf {
        let path = std::env::temp_dir().join(format!("masteredit-office-{name}.xlsx"));
        let _ = std::fs::remove_file(&path);
        let mut wb = Workbook::new();
        {
            let sheet = wb.add_worksheet();
            sheet.set_name("数据").unwrap();
            sheet.write_string(0, 0, "名称").unwrap();
            sheet.write_string(0, 1, "数量").unwrap();
            sheet.write_string(1, 0, "苹果").unwrap();
            sheet.write_number(1, 1, 12.5).unwrap();
            sheet.write_boolean(2, 0, true).unwrap();
            sheet.write_formula(2, 1, "=SUM(B2:B2)").unwrap();
            sheet.write_number(3, 1, 1000.0).unwrap();
        }
        {
            let second = wb.add_worksheet();
            second.set_name("第二表").unwrap();
            second.write_string(0, 0, "x").unwrap();
        }
        wb.save(&path).unwrap();
        path
    }

    #[test]
    fn lists_sheets() {
        let path = make_workbook("info");
        let info = info_impl(path.to_str().unwrap()).expect("应能读取工作簿");
        let names: Vec<_> = info.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, vec!["数据", "第二表"]);
        assert!(info.size > 0);
    }

    #[test]
    fn reads_window_with_types_and_formula() {
        let path = make_workbook("rows");
        let cache = SheetCache::default();
        let win = rows_impl(&cache, path.to_str().unwrap(), "数据", 0, 100).expect("应能读取窗口");
        assert_eq!(win.rows, 4);
        assert_eq!(win.cols, 2);
        assert_eq!(win.start, 0);
        assert_eq!(win.cells.len(), 4);

        assert_eq!(win.cells[0][0].v, "名称");
        assert_eq!(win.cells[0][0].t, "text");
        assert_eq!(win.cells[1][1].v, "12.5");
        assert_eq!(win.cells[1][1].t, "number");
        assert_eq!(win.cells[2][0].v, "true");
        assert_eq!(win.cells[2][0].t, "bool");
        // 公式带前导 =，数值取工作簿缓存值
        assert_eq!(win.cells[2][1].f.as_deref(), Some("=SUM(B2:B2)"));
        assert_eq!(win.cells[3][1].v, "1000");
        assert_eq!(win.cells[3][1].t, "number");
        assert!(!win.truncated);
    }

    /// 已用区域不从 A1 开始时，单元格仍要落在与 Excel 一致的绝对坐标上
    #[test]
    fn non_a1_ranges_keep_absolute_coordinates() {
        let path = std::env::temp_dir().join("masteredit-office-offset.xlsx");
        let _ = std::fs::remove_file(&path);
        let mut wb = Workbook::new();
        {
            let sheet = wb.add_worksheet();
            sheet.write_string(2, 1, "偏移").unwrap(); // B3
            sheet.write_number(3, 2, 7.0).unwrap(); // C4
        }
        wb.save(&path).unwrap();

        let cache = SheetCache::default();
        let win = rows_impl(&cache, path.to_str().unwrap(), "Sheet1", 0, 20).expect("应能读取");
        // 无论 calamine 的 Range 是否从 A1 起算，取值坐标都必须对齐
        assert_eq!(win.cells[2][1].v, "偏移");
        assert_eq!(win.cells[3][2].v, "7");
        assert_eq!(win.cells[0][0].t, "empty");
        assert!(win.rows >= 4);
        assert!(win.cols >= 3);
        // 越界窗口应被收敛到总行数
        let tail = rows_impl(&cache, path.to_str().unwrap(), "Sheet1", win.rows - 1, 20).unwrap();
        assert_eq!(tail.cells.len(), 1);
    }

    #[test]
    fn window_start_beyond_end_is_clamped() {
        let path = make_workbook("clamp");
        let cache = SheetCache::default();
        let win = rows_impl(&cache, path.to_str().unwrap(), "数据", 999, 50).expect("越界起点应被收敛");
        assert!(win.cells.is_empty());
        assert_eq!(win.start, win.rows);
    }

    #[test]
    fn unknown_sheet_reports_error() {
        let path = make_workbook("unknown");
        let cache = SheetCache::default();
        let err = rows_impl(&cache, path.to_str().unwrap(), "不存在", 0, 10).expect_err("应报错");
        assert!(err.contains("不存在"), "错误信息应包含表名: {err}");
    }

    #[test]
    fn caches_parsed_sheet_across_windows() {
        let path = make_workbook("cache");
        let cache = SheetCache::default();
        let p = path.to_str().unwrap();
        rows_impl(&cache, p, "数据", 0, 10).expect("首次读取");
        assert_eq!(cache.0.lock().unwrap().entries.len(), 1);
        rows_impl(&cache, p, "数据", 2, 10).expect("第二次读取应命中缓存");
        assert_eq!(cache.0.lock().unwrap().entries.len(), 1, "同一张表不应重复解析");
        rows_impl(&cache, p, "第二表", 0, 10).expect("另一张表");
        assert_eq!(cache.0.lock().unwrap().entries.len(), 2);
    }

    /// 真实样本（本机存在时才跑）：企业透明加密的 xlsx。
    /// 这类文件直接交给 calamine 会报 "Could not find EOCD"，必须先在内存中解密。
    #[test]
    fn real_esafenet_workbook_if_present() {
        let path = r"Z:\D\mywork\01_project\007_topology_identification\06_data\voltage_daily_v3\2026-03-01.xlsx";
        if !Path::new(path).exists() {
            return;
        }
        let info = info_impl(path).expect("真实加密工作簿应能打开");
        assert!(info.encrypted, "应识别为企业加密文档");
        assert!(!info.sheets.is_empty(), "应至少有一个工作表");

        let cache = SheetCache::default();
        let sheet = info.sheets[0].name.clone();
        let win = rows_impl(&cache, path, &sheet, 0, 20).expect("应能读取首个工作表");
        assert!(win.rows > 0, "行数应大于 0");
        assert!(win.cols > 0, "列数应大于 0");
        assert!(!win.cells.is_empty(), "首批窗口应有数据");
        println!(
            "真实加密样本：工作表 {:?} 共 {} 行 × {} 列，首格 {:?}",
            sheet, win.rows, win.cols, win.cells[0][0].v
        );
    }

    #[test]
    fn missing_file_reports_error() {
        let err = info_impl(r"C:\definitely\not\here.xlsx").expect_err("应报错");
        assert!(err.contains("读取文件信息失败"));
    }

    /// 回归：企业透明加密（亿赛通）的 xlsx —— 真实场景里 calamine 直接读原始字节会报
    /// "Could not find EOCD"，必须在内存中解密后再解析。
    #[test]
    fn decrypts_esafenet_wrapped_workbook() {
        let plain_path = make_workbook("wrapped");
        let plain = std::fs::read(&plain_path).expect("读取测试工作簿");
        // 与真实加密文件一致的头：魔数 + 块大小 512 + 文件头长度 4096
        let mut header = vec![0u8; 4096];
        header[0..4].copy_from_slice(&[0xE0, 0xA8, 0x91, 0xE7]);
        header[8..12].copy_from_slice(&512u32.to_le_bytes());
        header[12..16].copy_from_slice(&4096u32.to_le_bytes());
        let wrapped = esafenet::encrypt_esafenet(&header, &plain);

        let enc_path = std::env::temp_dir().join("masteredit-office-encrypted.xlsx");
        std::fs::write(&enc_path, &wrapped).expect("写入加密文件");
        let enc = enc_path.to_str().unwrap();

        let info = info_impl(enc).expect("加密工作簿应能解密后读取");
        assert!(info.encrypted, "应标记为企业加密文档");
        assert_eq!(info.sheets.len(), 2);

        let cache = SheetCache::default();
        let win = rows_impl(&cache, enc, "数据", 0, 50).expect("加密工作表应能解密后读取");
        assert_eq!(win.cells[0][0].v, "名称");
        assert_eq!(win.cells[1][1].v, "12.5");
        assert_eq!(win.cells[2][1].f.as_deref(), Some("=SUM(B2:B2)"));
    }

    /// 需要密码的 Office 文档 / 旧版二进制格式都是 CFB 容器：提示要可执行，而不是抛 zip 错误
    #[test]
    fn gives_actionable_error_for_cfb_container() {
        let path = std::env::temp_dir().join("masteredit-office-cfb.xlsx");
        let mut bytes = vec![0xD0, 0xCF, 0x11, 0xE0];
        bytes.extend_from_slice(&[0u8; 508]);
        std::fs::write(&path, &bytes).unwrap();
        let err = info_impl(path.to_str().unwrap()).expect_err("应报错");
        assert!(err.contains("加密的 Office 文档"), "实际提示：{err}");
    }

    /// 系统导出的「假 Excel」（真实内容是 HTML 表格）要单独提示
    #[test]
    fn gives_actionable_error_for_html_export() {
        let path = std::env::temp_dir().join("masteredit-office-html.xlsx");
        std::fs::write(
            &path,
            b"<html><head><meta charset=\"utf-8\"></head><body><table><tr><td>1</td></tr></table></body></html>",
        )
        .unwrap();
        let err = info_impl(path.to_str().unwrap()).expect_err("应报错");
        assert!(err.contains("网页表格"), "实际提示：{err}");
    }

    #[test]
    fn excel_serial_dates_match_expected() {
        // 期望值已用 .NET DateTime 逐点核对（见函数注释）
        assert_eq!(format_serial(1.0, false), "1900-01-01");
        assert_eq!(format_serial(59.0, false), "1900-02-28");
        assert_eq!(format_serial(61.0, false), "1900-03-01");
        assert_eq!(format_serial(25569.0, false), "1970-01-01");
        assert_eq!(format_serial(44927.0, false), "2023-01-01");
        assert_eq!(format_serial(45658.0, false), "2025-01-01");
        assert_eq!(format_serial(45292.5, false), "2024-01-01 12:00:00");
        // 纯时间与时长
        assert_eq!(format_serial(0.5, false), "12:00:00");
        assert_eq!(format_serial(1.5, true), "36:00:00");
        // 非有限值原样返回，不 panic
        assert_eq!(format_serial(f64::INFINITY, false), "inf");
    }

    #[test]
    fn float_formatting_is_compact() {
        assert_eq!(format_float(1000.0), "1000");
        assert_eq!(format_float(12.5), "12.5");
        assert_eq!(format_float(-0.0), "0");
        assert_eq!(format_float(0.1), "0.1");
    }

    #[test]
    fn long_text_is_truncated_on_char_boundary() {
        let long = "中".repeat(MAX_CELL_CHARS + 50);
        let out = truncate_text(&long);
        assert_eq!(out.chars().count(), MAX_CELL_CHARS + 1);
        assert!(out.ends_with('…'));
    }
}
