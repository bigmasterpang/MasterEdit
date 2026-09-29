//! xlsx 单元格写回（0.19.0「表格轻量编辑」后端）。
//!
//! 读取仍由 `office.rs`（calamine，只读）负责；本模块只在**保存**时用
//! umya-spreadsheet 3.1 做「读入 → 改值 → 写出」的整簿重写。
//!
//! 设计要点：
//! - **只有 `.xlsx` 可写回**：xlsm 含 VBA 宏、xls / xlsb / ods 结构完全不同，
//!   整簿重写会把它们丢掉，因此一律拒绝（前端据 `SpreadsheetInfo.editable` 判断）。
//! - **整簿重写**：umya 不支持局部改写 zip，保存即重新生成整个工作簿。未编辑的
//!   工作表、样式、公式都会带过去（`cell_mut` 拿到的就是文件里原有的 Cell，样式对象
//!   不动），但内部结构会被 umya 重排 —— 这也是默认先备份 `<文件名>.bak` 的原因。
//! - **公式不重算**：umya 没有 calcPr / fullCalcOnLoad 接口，本模块**不**去强制重算。
//!   用户写入的公式只写公式文本（缓存值保持原样），其它公式单元格的缓存值同样原样
//!   保留（可能过期），需要时由 Excel / WPS 打开后自行计算。
//! - **企业透明加密**：解密、加密全程在内存里完成，**不产生明文临时文件**；写盘前用
//!   calamine 校验**明文**可读，落盘时再用原文件的 4096 字节头重新加密。
//! - **原子替换**：先写目标同目录的临时文件，回读校验后用 `fs::rename` 覆盖目标；
//!   任何一步失败都会清理临时文件并让原文件保持不变。

use crate::commands::esafenet;
use crate::commands::file::modified_ms;
use calamine::{open_workbook_auto_from_rs, Reader};
use std::io::Cursor;
use std::path::{Path, PathBuf};

/// xlsx 规格上限（Excel 2010+）：越界直接报错，避免把非法坐标交给 umya
const MAX_ROWS: u32 = 1_048_576;
const MAX_COLS: u32 = 16_384;

/// 临时文件前缀（与目标同目录，保存失败时清理，正常保存后由 rename 消耗掉）
const TMP_PREFIX: &str = ".MasterEdit-tmp-";

/// 前端提交的一处单元格改动（坐标为 0 起的绝对行列号，与网格一致）
#[derive(serde::Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct CellEdit {
    pub sheet: String,
    /// 0 起的绝对行号（与前端网格坐标一致，Rust 侧 +1 后交给 umya）
    pub row: u32,
    /// 0 起的绝对列号
    pub col: u32,
    /// "text" | "number" | "bool" | "formula" | "date" | "empty"
    pub kind: String,
    pub value: String,
}

/// 保存结果：前端据此刷新基线（大小 / 修改时间）与提示备份位置
#[derive(serde::Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct SaveResult {
    pub path: String,
    pub size: u64,
    pub modified_at: u64,
    pub saved_cells: usize,
    /// 备份文件路径（未备份时为 None）
    pub backup_path: Option<String>,
}

/// 是否支持写回编辑：只有 .xlsx 可以（xlsm 含宏、xls/xlsb/ods 结构不同，一律只读）。
/// `office.rs` 的 `SpreadsheetInfo.editable` 直接用它，保证前后端判断一致。
pub fn is_editable(path: &str) -> bool {
    Path::new(path)
        .extension()
        .map(|e| e.eq_ignore_ascii_case("xlsx"))
        .unwrap_or(false)
}

/// 不可写回时的统一报错（保存流水线与结构操作共用同一条话术）
pub(crate) fn editable_required(path: &str) -> Result<(), String> {
    if is_editable(path) {
        Ok(())
    } else {
        Err(
            "只有 .xlsx 支持写回编辑（xlsm 含宏、xls/xlsb/ods 结构不同），请另存为 .xlsx 后再编辑"
                .to_string(),
        )
    }
}

/* ------------------------------------------------------------------ */
/* 日期换算：文本 → Excel 序列号（office.rs::format_serial 的反向）      */
/* ------------------------------------------------------------------ */

/// Unix 天数 → (年, 月, 日)：Howard Hinnant 的 civil_from_days（与 image.rs 同算法）。
/// 这里只用于「日期文本是否真实存在」的往返校验。
fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = (z - era * 146_097) as u64;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
}

/// (年, 月, 日) → Unix 天数：Hinnant 的 days_from_civil。
/// 调用前必须保证月份 1..=12、日 1..=31（否则无符号运算会下溢）。
fn days_from_civil(y: i64, m: u32, d: u32) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400; // [0, 399]
    let mp = ((m + 9) % 12) as i64; // [0, 11]
    let doy = (153 * mp + 2) / 5 + d as i64 - 1; // [0, 365]
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy; // [0, 146096]
    era * 146_097 + doe - 719_468
}

/// 解析 `YYYY-MM-DD` / `YYYY/M/D`，并校验日期真实存在（2024-02-31 之类直接判为非法）。
fn parse_date_parts(part: &str) -> Option<(i64, u32, u32)> {
    let part = part.trim();
    let sep = if part.contains('-') {
        '-'
    } else if part.contains('/') {
        '/'
    } else {
        return None;
    };
    let mut it = part.split(sep);
    let y: i64 = it.next()?.trim().parse().ok()?;
    let m: u32 = it.next()?.trim().parse().ok()?;
    let d: u32 = it.next()?.trim().parse().ok()?;
    if it.next().is_some() || !(1..=12).contains(&m) || !(1..=31).contains(&d) {
        return None;
    }
    // 往返校验：把天数再拆回日期，能对上才说明这一天真实存在
    if civil_from_days(days_from_civil(y, m, d)) != (y, m, d) {
        return None;
    }
    Some((y, m, d))
}

/// 解析 `HH:MM` / `HH:MM:SS` → 一天内的秒数
fn parse_time_secs(part: &str) -> Option<f64> {
    let part = part.trim();
    if part.is_empty() {
        return None;
    }
    let mut it = part.split(':');
    let h: u32 = it.next()?.trim().parse().ok()?;
    let mi: u32 = it.next()?.trim().parse().ok()?;
    let sec: f64 = match it.next() {
        Some(s) => s.trim().parse().ok()?,
        None => 0.0,
    };
    if it.next().is_some() || h > 23 || mi > 59 || !(0.0..60.0).contains(&sec) {
        return None;
    }
    Some(h as f64 * 3600.0 + mi as f64 * 60.0 + sec)
}

/// 日期文本 → Excel 序列号（1900 日期系统）。
///
/// 支持 `YYYY-MM-DD`、`YYYY/M/D`、`YYYY-MM-DD HH:MM[:SS]`、`YYYY/M/D H:MM`（也容忍
/// 用 `T` 分隔日期与时间）。与 `office.rs::format_serial` 互为反函数，锚点已逐点核对：
/// 1900-01-01→1、1900-02-28→59、1900-03-01→61、1970-01-01→25569、2023-01-01→44927、
/// 2024-01-01→45292。**1900-01-01 之前（序列号 < 1）不解析**，返回 None 由调用方按
/// 字符串写入，避免用户看到 0 或负数这种怪值。
fn serial_from_text(text: &str) -> Option<f64> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }
    let (date_part, time_part) = match text.split_once(' ').or_else(|| text.split_once('T')) {
        Some((d, t)) => (d, Some(t)),
        None => (text, None),
    };
    let (y, m, d) = parse_date_parts(date_part)?;
    let days = days_from_civil(y, m, d);
    // 反向映射：1900-03-01（序列号 61）起以 1899-12-30 为原点，之前的日期需回退一天
    let serial = if days >= -25_508 {
        days + 25_569
    } else {
        days + 25_568
    };
    if serial < 1 {
        return None;
    }
    let mut value = serial as f64;
    if let Some(t) = time_part {
        value += parse_time_secs(t)? / 86_400.0;
    }
    Some(value)
}

/// 数字格式码是否为「日期 / 时间」格式。
///
/// 判断方式：先去掉引号里的字面量、方括号段（`[Red]`、`[h]` 这类颜色/条件）、
/// 转义字符（`\x`、`_x`、`*x`），再看剩下的格式码里有没有 y / m / d / h 这些
/// 日期时间占位符。这样 `0.00" m"`（单位是米）不会被误判成日期，而
/// `m/d/yyyy`、`h:mm:ss`、`[h]:mm:ss`、`yyyy"年"m"月"` 都能识别；纯文本格式 `@`
/// 与 `General` 自然落在「非日期」一侧。
fn format_code_is_date(code: &str) -> bool {
    let mut cleaned = String::with_capacity(code.len());
    let mut quoted = false;
    let mut chars = code.chars();
    while let Some(c) = chars.next() {
        match c {
            '"' => quoted = !quoted,
            '[' if !quoted => {
                for n in chars.by_ref() {
                    if n == ']' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' if !quoted => {
                let _ = chars.next();
            }
            _ => {
                if !quoted {
                    cleaned.push(c);
                }
            }
        }
    }
    cleaned
        .to_ascii_lowercase()
        .chars()
        .any(|c| matches!(c, 'y' | 'm' | 'd' | 'h'))
}

/* ------------------------------------------------------------------ */
/* 编辑应用                                                            */
/* ------------------------------------------------------------------ */

/// 清掉单元格上的公式对象。
///
/// 注意：**不能**用 `set_formula("")` —— umya 3.1 里它只是塞进一个「空公式对象」，
/// 写出的 XML 会留下 `<f></f>`（Excel 打开时可能提示修复）。真正清除公式的是
/// `CellValue::remove_formula()`；`set_value_string / set_value_number / set_value_bool /
/// set_blank` 内部也已经调用了它，这里再显式调一次，语义更明确。
fn clear_formula(cell: &mut umya_spreadsheet::Cell) {
    cell.cell_value_mut().remove_formula();
}

/// 数字文本 → f64（拒绝空值与非有限值，避免把 NaN / inf 写进文件）。
///
/// `pub(crate)`：`office_ops::spreadsheet_stats` 选区统计里「文本型的数字」（CSV 导入后
/// 很常见）按同一套语义判定，保证「统计里算进的数字」与「能写回的数字」一致。
pub(crate) fn parse_number(text: &str) -> Option<f64> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }
    let value: f64 = text.parse().ok()?;
    if value.is_finite() {
        Some(value)
    } else {
        None
    }
}

/// 布尔文本 → bool：接受 true/false/TRUE/FALSE/1/0。
/// `pub(crate)`：公式求值把 `bool` 类型的单元格编辑转成值时复用同一份口径。
pub(crate) fn parse_bool(text: &str) -> Option<bool> {
    match text.trim().to_ascii_lowercase().as_str() {
        "true" | "1" => Some(true),
        "false" | "0" => Some(false),
        _ => None,
    }
}

/// 坐标越界（xlsx 规格）直接拒绝：umya 坐标 1 起，`u32::MAX` 这类值 +1 会溢出
fn check_bounds(edit: &CellEdit) -> Result<(), String> {
    if edit.row >= MAX_ROWS || edit.col >= MAX_COLS {
        return Err(format!(
            "单元格坐标越界：第 {} 行第 {} 列（xlsx 上限 {MAX_ROWS} 行 × {MAX_COLS} 列）",
            edit.row as u64 + 1,
            edit.col as u64 + 1
        ));
    }
    Ok(())
}

/// 把编辑应用到内存中的工作簿，返回应用条数。
///
/// 全部在内存里完成：任何一处失败（表名不存在、数字解析失败……）都直接返回 Err，
/// 调用方不会落盘，原文件分毫不动。
///
/// `pub(crate)`：结构操作（`office_ops::spreadsheet_structure`）要先应用同一批编辑再改结构，
/// 走的必须是这一份实现，避免两处对「kind / 日期 / 公式」的解释出现分歧。
pub(crate) fn apply_edits(
    workbook: &mut umya_spreadsheet::Workbook,
    edits: &[CellEdit],
) -> Result<usize, String> {
    let mut saved = 0usize;
    for edit in edits {
        check_bounds(edit)?;
        let sheet = workbook.sheet_by_name_mut(&edit.sheet).map_err(|_| {
            format!(
                "工作簿里没有工作表「{}」，无法写入第 {} 行第 {} 列",
                edit.sheet,
                edit.row as u64 + 1,
                edit.col as u64 + 1
            )
        })?;
        // umya 坐标 1 起，元组是 (列, 行)：+1 换到它的绝对坐标
        let cell = sheet.cell_mut((edit.col + 1, edit.row + 1));
        let at = format!(
            "工作表「{}」第 {} 行第 {} 列",
            edit.sheet,
            edit.row as u64 + 1,
            edit.col as u64 + 1
        );
        match edit.kind.as_str() {
            "text" => {
                cell.set_value_string(edit.value.as_str());
                clear_formula(cell);
            }
            "number" => {
                let value = parse_number(&edit.value).ok_or_else(|| {
                    format!("{at}：无法把「{}」解析成数字", edit.value.trim())
                })?;
                cell.set_value_number(value);
                clear_formula(cell);
            }
            "bool" => {
                let value = parse_bool(&edit.value).ok_or_else(|| {
                    format!("{at}：无法把「{}」解析成布尔值（可用 true/false/1/0）", edit.value.trim())
                })?;
                cell.set_value_bool(value);
                clear_formula(cell);
            }
            "formula" => {
                let text = edit.value.trim();
                let formula = text.strip_prefix('=').unwrap_or(text).trim();
                if formula.is_empty() {
                    return Err(format!("{at}：公式不能为空"));
                }
                // 公式在 umya 里不带前导 =；缓存值保持原样（不重算，见模块注释）
                cell.set_formula(formula);
            }
            "date" => {
                // 目标格是日期格式才写序列号（这样仍按日期显示）；否则把原文当字符串，
                // 免得用户看到 45292 这种数字。样式全程不动。
                let date_style = {
                    let code = cell
                        .style()
                        .number_format()
                        .map(|f| f.format_code())
                        .unwrap_or("");
                    format_code_is_date(code)
                };
                match if date_style { serial_from_text(&edit.value) } else { None } {
                    Some(serial) => {
                        cell.set_value_number(serial);
                    }
                    None => {
                        cell.set_value_string(edit.value.as_str());
                    }
                }
                clear_formula(cell);
            }
            "empty" => {
                cell.set_blank();
                clear_formula(cell);
            }
            other => {
                return Err(format!(
                    "{at}：不支持的单元格类型「{other}」（可用 text/number/bool/formula/date/empty）"
                ))
            }
        }
        saved += 1;
    }
    Ok(saved)
}

/* ------------------------------------------------------------------ */
/* 读写与落盘                                                          */
/* ------------------------------------------------------------------ */

/// calamine 打不开明文时的提示：区分「根本不是 xlsx」与「xlsx 损坏」。
/// `pub(crate)`：影子工作簿首次从磁盘读入时复用同一条提示。
pub(crate) fn describe_open_failure(bytes: &[u8], err: &str) -> String {
    if bytes.len() >= 4 && bytes[..4] == [0xD0, 0xCF, 0x11, 0xE0] {
        return format!(
            "该文件不是 xlsx（是旧版 .xls、需要密码的 Office 文档或其它 CFB 容器），无法写回编辑。\n\
             请先用 Excel / WPS 另存为 .xlsx。（底层错误：{err}）"
        );
    }
    if bytes.len() < 2 || bytes[..2] != *b"PK" {
        return format!(
            "该文件的真实格式不是 xlsx（扩展名与内容不一致），无法写回编辑。\n（底层错误：{err}）"
        );
    }
    format!("无法解析工作簿，文件可能已损坏或被其它程序占用。\n（底层错误：{err}）")
}

/// 用 calamine 从内存校验输出：能解析、且工作表名齐全
fn validate_output(plain: &[u8], expected: &[String]) -> Result<(), String> {
    let workbook = open_workbook_auto_from_rs(Cursor::new(plain))
        .map_err(|e| format!("保存校验失败：生成的 xlsx 无法解析（{e}）"))?;
    let names = workbook.sheet_names();
    for want in expected {
        if !names.iter().any(|n| n == want) {
            return Err(format!("保存校验失败：生成的 xlsx 缺少工作表「{want}」"));
        }
    }
    Ok(())
}

/// 两个路径是否指向同一个文件（target 与原路径相同时按普通保存处理）
fn same_file(source: &Path, dest: &Path) -> bool {
    if source == dest {
        return true;
    }
    match (source.canonicalize(), dest.canonicalize()) {
        (Ok(a), Ok(b)) => a == b,
        _ => false,
    }
}

/// 解析写回目标：None / 空串 / 与原路径相同 → 写回原文件。
/// 另存为时目录必须已存在（不自动建目录，免得用户手滑在奇怪的位置建出一堆目录）。
///
/// `pub(crate)`：影子保存（`office_shadow::save_with_shadow`）也要走同一套目标解析。
pub(crate) fn resolve_target(source: &Path, target: Option<&str>) -> Result<PathBuf, String> {
    let Some(text) = target.map(str::trim).filter(|t| !t.is_empty()) else {
        return Ok(source.to_path_buf());
    };
    let dest = PathBuf::from(text);
    if same_file(source, &dest) {
        return Ok(source.to_path_buf());
    }
    if dest.is_dir() {
        return Err(format!("另存为的目标是文件夹：{}", dest.display()));
    }
    if let Some(dir) = dest.parent().filter(|p| !p.as_os_str().is_empty()) {
        if !dir.is_dir() {
            return Err(format!("另存为的目录不存在：{}", dir.display()));
        }
    }
    // 只允许 .xlsx：写出「名字是 .xls、内容其实是 xlsx」的假文件比直接报错糟糕得多
    match dest.extension().and_then(|e| e.to_str()) {
        Some(ext) if ext.eq_ignore_ascii_case("xlsx") => Ok(dest),
        Some(ext) => Err(format!("只能另存为 .xlsx，不能保存为 .{ext}")),
        None => Ok(dest.with_extension("xlsx")),
    }
}

/// 备份原文件的**原始字节**（加密文件就是加密后的密文，绝不解密）为 `<文件名>.bak`
fn write_backup(source: &Path, raw: &[u8]) -> Result<PathBuf, String> {
    let name = source
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .ok_or_else(|| "无法确定原文件名，备份失败".to_string())?;
    let dir = source
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));
    let backup = dir.join(format!("{name}.bak"));
    std::fs::write(&backup, raw).map_err(|e| format!("写入备份文件失败: {e}"))?;
    Ok(backup)
}

/// 目标同目录的临时文件路径：`.MasterEdit-tmp-<pid>-<文件名>`
fn temp_path_for(target: &Path) -> PathBuf {
    let name = target
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_else(|| "out.xlsx".to_string());
    let dir = target
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));
    dir.join(format!("{TMP_PREFIX}{}-{name}", std::process::id()))
}

/// 原子替换：写临时文件 → 回读校验 → `fs::rename` 覆盖目标。
///
/// `expected_plain` 传加密文件的**明文**（回读后解密比对，确认重新加密没写坏）；
/// 非加密文件传 None，内容正确性由调用方的 calamine 内存校验保证。任何失败都会
/// 删除临时文件，原文件保持不变。
fn atomic_write(dest: &Path, bytes: &[u8], expected_plain: Option<&[u8]>) -> Result<(), String> {
    let tmp = temp_path_for(dest);
    let result = (|| -> Result<(), String> {
        use std::io::Write;
        let mut file =
            std::fs::File::create(&tmp).map_err(|e| format!("创建临时文件失败: {e}"))?;
        file.write_all(bytes)
            .map_err(|e| format!("写入临时文件失败: {e}"))?;
        // 尽力而为地刷盘：rename 之后即便掉电，也不至于留下半截文件
        let _ = file.sync_all();
        drop(file);

        let written = std::fs::read(&tmp).map_err(|e| format!("校验临时文件失败: {e}"))?;
        if written != bytes {
            return Err("校验临时文件失败：回读内容与写入内容不一致".to_string());
        }
        if let Some(plain) = expected_plain {
            match esafenet::decrypt_esafenet(&written) {
                Some(back) if back == plain => {}
                _ => return Err("校验临时文件失败：重新加密的内容无法还原".to_string()),
            }
        }
        // Windows 上 Rust 的 rename 会替换已存在文件（MoveFileEx + REPLACE_EXISTING）
        std::fs::rename(&tmp, dest).map_err(|e| format!("替换目标文件失败: {e}"))?;
        Ok(())
    })();
    if result.is_err() {
        let _ = std::fs::remove_file(&tmp);
    }
    result
}

/// 序列化工作簿为 xlsx 字节。
///
/// 写盘尾段与影子工作簿的 `serialized` 缓存共用这一份实现，保证「读到的影子内容」与
/// 「将来落盘的内容」逐字节一致。
pub(crate) fn serialize_workbook(workbook: &umya_spreadsheet::Workbook) -> Result<Vec<u8>, String> {
    let mut out: Vec<u8> = Vec::new();
    umya_spreadsheet::writer::xlsx::write_writer(workbook, &mut out)
        .map_err(|e| format!("生成 xlsx 失败: {e}"))?;
    Ok(out)
}

/// 写盘尾段（同步）：序列化 → calamine 校验 → （加密文件）用原头重新加密 → 备份 → 原子替换。
///
/// - `raw` 是**原文件的原始字节**（企业加密文件即密文）：既是备份内容，也是重新加密时
///   沿用的 4096 字节头来源；
/// - `encrypted` 为 true 时写出的是密文，并额外校验「回读后能解密回明文」；
/// - `backup_source` 为 Some 时把 `raw` 复制成 `<该文件名>.bak`（另存为时传源文件路径）。
///
/// 与 `save_pipeline` 里的第 4~8 步完全一致，抽出来是为了让「影子工作簿」也能落盘而
/// 不必重新走一遍磁盘读取。
pub(crate) fn write_back_workbook(
    dest: &Path,
    workbook: &umya_spreadsheet::Workbook,
    saved_cells: usize,
    backup_source: Option<&Path>,
    raw: &[u8],
    encrypted: bool,
) -> Result<SaveResult, String> {
    // 4) 写回内存
    let mut out = serialize_workbook(workbook)?;

    // 5) 用 calamine 从内存校验明文（加密文件也校验加密前的明文）
    let expected: Vec<String> = workbook
        .sheet_collection_no_check()
        .iter()
        .map(|s| s.name().to_string())
        .collect();
    validate_output(&out, &expected)?;

    // 6) 加密文件重新加密：沿用原文件的 4096 字节头
    let payload = if encrypted {
        let head_len = raw.len().min(4096);
        esafenet::encrypt_esafenet(&raw[..head_len], &out)
    } else {
        std::mem::take(&mut out)
    };

    // 7) 备份原文件的原始字节（另存为时备份的是源文件），失败即中止保存
    let backup_path = match backup_source {
        Some(source) => Some(write_backup(source, raw)?),
        None => None,
    };

    // 8) 原子替换
    atomic_write(
        dest,
        &payload,
        if encrypted { Some(out.as_slice()) } else { None },
    )?;

    let meta = std::fs::metadata(dest).map_err(|e| format!("读取文件信息失败: {e}"))?;
    Ok(SaveResult {
        path: dest.to_string_lossy().to_string(),
        size: meta.len(),
        modified_at: modified_ms(&meta),
        saved_cells,
        backup_path: backup_path.map(|p| p.to_string_lossy().to_string()),
    })
}

/// 保存流水线（同步）：读入工作簿 → 由 `mutate` 应用改动 → 写出 → 校验 → 重新加密
/// → 备份 → 原子替换。
///
/// `mutate` 拿到**已在内存里解密并解析好**的工作簿，返回本次应用的编辑格数；返回 Err
/// 时整个流程中止，磁盘上的文件分毫未动（也不会留下备份或临时文件）。两处调用：
/// - `office_shadow::save`：该文件**没有**内存影子时走这里（行为与 0.19.0 完全一致，
///   只逐条应用单元格编辑）；有影子时它改用影子工作簿 + [`write_back_workbook`]；
/// - 本模块的单元测试：直接验证磁盘读写本身。
///
/// 只有 `.xlsx` 可写回；企业加密文件全程内存解密/加密，不落明文；备份的是原文件的
/// **原始字节**；失败一律清理临时文件并保持原文件不变。
pub(crate) fn save_pipeline(
    path: &str,
    target: Option<&str>,
    backup: bool,
    mutate: impl FnOnce(&mut umya_spreadsheet::Workbook) -> Result<usize, String>,
) -> Result<SaveResult, String> {
    let source = Path::new(path);
    // 只有 .xlsx 能整簿重写：xlsm 会被连宏一起丢掉，xls/xlsb/ods 直接解析失败
    editable_required(path)?;
    let meta = std::fs::metadata(source).map_err(|e| format!("读取文件信息失败: {e}"))?;
    if !meta.is_file() {
        return Err("目标不是有效的文件".to_string());
    }
    let dest = resolve_target(source, target)?;

    // 1) 读原始字节：密文原始字节要用于备份与重新加密的文件头
    let raw = std::fs::read(source).map_err(|e| format!("读取文件失败: {e}"))?;
    let encrypted = esafenet::is_esafenet_encrypted(&raw);
    let plain = if encrypted {
        // 企业透明加密：内存解密，绝不落明文临时文件
        esafenet::decrypt_esafenet(&raw).ok_or_else(|| "企业加密文档解密失败".to_string())?
    } else {
        raw.clone()
    };

    // 2) umya 读入（内存）
    let mut workbook =
        umya_spreadsheet::reader::xlsx::read_reader(Cursor::new(plain.as_slice()), true)
            .map_err(|e| describe_open_failure(&plain, &e.to_string()))?;

    // 3) 应用改动：失败即返回，磁盘上的文件还没被碰过
    let saved_cells = mutate(&mut workbook)?;

    // 4~8) 写出 → 校验 → 重新加密 → 备份 → 原子替换
    write_back_workbook(
        &dest,
        &workbook,
        saved_cells,
        if backup { Some(source) } else { None },
        &raw,
        encrypted,
    )
}

/// 把前端网格里的编辑写回 xlsx（0.19.0「表格轻量编辑」）。
///
/// - **有内存影子就用影子**：结构操作（`spreadsheet_structure`）只改内存，保存时才在这里
///   把影子（加上本次 `edits`）整簿写回；写回成功后影子与撤销栈都保留（保存不影响撤回），
///   只把「磁盘基线」挪到当前状态。没有影子时完全保持原有行为：读盘 → 应用编辑 → 写出。
/// - `target` 为 None 时写回原文件；`Some(路径)` 为另存为（此时备份的是**源文件**，
///   且原文件的影子保持脏状态 —— 它仍有未落盘的改动）；
/// - `backup` **默认 false：不再生成 `.bak`**（0.20.0 起），只有显式传 true 才把改动前的
///   原始字节复制为 `<文件名>.bak`；企业加密文件备份的也是密文，不落明文；
/// - 只有 `.xlsx` 可写回，其它格式（xlsm/xls/xlsb/ods）一律报错；
/// - 保存是**整簿重写**：先用 calamine 从内存校验输出可读、工作表名齐全，再写同目录
///   临时文件、回读校验后用 `fs::rename` 原子替换目标；失败会清理临时文件并保持原文件不变；
/// - **公式不重算**：只写用户输入。新建/修改的公式只写公式文本（缓存值保持不变），
///   其它公式单元格的缓存值原样带过去，可能过期 —— 由 Excel / WPS 打开后自行计算。
#[tauri::command]
pub async fn spreadsheet_save(
    path: String,
    target: Option<String>,
    edits: Vec<CellEdit>,
    backup: Option<bool>,
) -> Result<SaveResult, String> {
    // 不传 = 不备份（0.20.0 起不再默认生成 .bak）
    tauri::async_runtime::spawn_blocking(move || {
        crate::commands::office_shadow::save_requested(
            &crate::commands::office_shadow::SHADOWS,
            &path,
            target.as_deref(),
            &edits,
            backup,
        )
    })
    .await
    .map_err(|e| format!("保存表格失败: {e}"))?
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::{Data, Sheets};
    use rust_xlsxwriter::{Color, Format, Workbook as XlsxWriter};
    use umya_spreadsheet::Style;

    /// 每个测试一个独立目录，避免并行执行时互相干扰（也方便断言目录里没有残留文件）
    fn test_dir(name: &str) -> PathBuf {
        let dir = std::env::temp_dir()
            .join("masteredit-office-write")
            .join(name);
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).expect("创建测试目录");
        dir
    }

    /// 样本：两个工作表；文本 / 数字 / 布尔 / 公式；B2 带自定义样式；B5 是日期格式
    fn make_sample(path: &Path) {
        let mut wb = XlsxWriter::new();
        let styled = Format::new().set_bold().set_background_color(Color::Yellow);
        let date_format = Format::new().set_num_format("yyyy-mm-dd");
        {
            let sheet = wb.add_worksheet();
            sheet.set_name("数据").unwrap();
            sheet.write_string(0, 0, "名称").unwrap();
            sheet.write_string(0, 1, "数量").unwrap();
            sheet.write_string(0, 2, "备注").unwrap();
            sheet.write_string(1, 0, "苹果").unwrap();
            // B2：带自定义样式（加粗 + 黄底）的数字文本单元格
            sheet.write_string_with_format(1, 1, "12.5", &styled).unwrap();
            // C2：不会被编辑的公式，用来验证「未编辑的公式仍在」
            sheet.write_formula(1, 2, "=SUM(B2:B2)").unwrap();
            sheet.write_boolean(2, 0, true).unwrap();
            sheet.write_number(2, 1, 3.0).unwrap();
            sheet.write_string(3, 0, "合计").unwrap();
            sheet.write_formula(3, 1, "=SUM(B2:B3)").unwrap();
            sheet.write_string(4, 0, "日期").unwrap();
            sheet.write_string_with_format(4, 1, "2024-01-01", &date_format).unwrap();
            sheet.write_string(5, 0, "纯文本").unwrap();
            sheet.write_string(5, 1, "旧").unwrap();
            sheet.write_string(5, 3, "x").unwrap();
        }
        {
            let second = wb.add_worksheet();
            second.set_name("第二表").unwrap();
            second.write_string(0, 0, "x").unwrap();
            second.write_number(1, 1, 7.0).unwrap();
        }
        wb.save(path).expect("生成样本 xlsx");
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

    fn save(
        path: &Path,
        target: Option<&str>,
        edits: &[CellEdit],
        backup: bool,
    ) -> Result<SaveResult, String> {
        // 直接走磁盘流水线（不带影子）：本模块的测试验证的就是磁盘读写本身，
        // 影子相关的行为在 office_shadow 的测试里
        save_pipeline(path.to_str().unwrap(), target, backup, |workbook| {
            apply_edits(workbook, edits)
        })
    }

    /// 用 calamine 重新打开文件（模拟前端读回）
    fn open(path: &Path) -> Sheets<Cursor<Vec<u8>>> {
        let bytes = std::fs::read(path).expect("读取保存后的文件");
        open_workbook_auto_from_rs(Cursor::new(bytes)).expect("保存后的文件应能被 calamine 解析")
    }

    /// 读一个单元格（绝对坐标，0 起）
    fn cell_at(wb: &mut Sheets<Cursor<Vec<u8>>>, sheet: &str, row: u32, col: u32) -> Data {
        wb.worksheet_range(sheet)
            .expect("工作表应存在")
            .get_value((row, col))
            .cloned()
            .unwrap_or(Data::Empty)
    }

    fn read_cell(path: &Path, sheet: &str, row: u32, col: u32) -> Data {
        cell_at(&mut open(path), sheet, row, col)
    }

    /// 读公式（calamine 的公式不含前导 =）；该格没有公式时返回 None
    fn read_formula(path: &Path, sheet: &str, row: u32, col: u32) -> Option<String> {
        let mut wb = open(path);
        let range = wb.worksheet_formula(sheet).ok()?;
        range
            .get_value((row, col))
            .map(|s| s.trim().trim_start_matches('=').to_string())
            .filter(|s| !s.is_empty())
    }

    /// 用 umya 读回单元格样式（验证编辑不动样式）
    fn style_at(path: &Path, sheet: &str, coord: (u32, u32)) -> Style {
        let bytes = std::fs::read(path).unwrap();
        let wb = umya_spreadsheet::reader::xlsx::read_reader(Cursor::new(bytes), true).unwrap();
        wb.sheet_by_name(sheet)
            .unwrap()
            .cell(coord)
            .map(|c| c.style().clone())
            .unwrap_or_default()
    }

    /// 用 umya 读回单元格数字格式码
    fn number_format_at(path: &Path, sheet: &str, coord: (u32, u32)) -> String {
        let bytes = std::fs::read(path).unwrap();
        let wb = umya_spreadsheet::reader::xlsx::read_reader(Cursor::new(bytes), true).unwrap();
        wb.sheet_by_name(sheet)
            .unwrap()
            .cell(coord)
            .and_then(|c| c.style().number_format())
            .map(|f| f.format_code().to_string())
            .unwrap_or_default()
    }

    /// 1) 四类编辑写回后：新值正确、未编辑的公式仍在、另一个工作表未受影响
    #[test]
    fn applies_edits_and_keeps_formulas_and_other_sheets() {
        let dir = test_dir("apply");
        let path = dir.join("sample.xlsx");
        make_sample(&path);

        let edits = [
            edit("数据", 1, 0, "text", "香蕉"),          // A2 文本
            edit("数据", 1, 1, "number", "42.5"),        // B2 数字（带自定义样式）
            edit("数据", 2, 1, "bool", "FALSE"),         // B3 布尔
            edit("数据", 3, 1, "formula", "=SUM(B2:B3)"), // B4 公式
        ];
        let result = save(&path, None, &edits, true).expect("保存应成功");
        assert_eq!(result.saved_cells, 4, "应记录 4 处编辑");
        assert_eq!(Path::new(&result.path), path.as_path());
        assert!(result.size > 0, "保存后大小应大于 0");
        assert!(result.modified_at > 0, "保存后应有修改时间");

        let mut wb = open(&path);
        assert_eq!(wb.sheet_names(), vec!["数据", "第二表"]);
        assert_eq!(cell_at(&mut wb, "数据", 1, 0), Data::String("香蕉".into()));
        assert_eq!(cell_at(&mut wb, "数据", 1, 1), Data::Float(42.5));
        assert_eq!(cell_at(&mut wb, "数据", 2, 1), Data::Bool(false));
        // 未编辑的表头与布尔单元格保持原值
        assert_eq!(cell_at(&mut wb, "数据", 0, 0), Data::String("名称".into()));
        assert_eq!(cell_at(&mut wb, "数据", 2, 0), Data::Bool(true));
        // 另一张工作表未受影响
        assert_eq!(cell_at(&mut wb, "第二表", 0, 0), Data::String("x".into()));
        assert_eq!(cell_at(&mut wb, "第二表", 1, 1), Data::Float(7.0));
        drop(wb);

        // 未编辑的公式（C2）仍在；被编辑成公式的 B4 也在
        assert_eq!(read_formula(&path, "数据", 1, 2).as_deref(), Some("SUM(B2:B2)"));
        assert_eq!(read_formula(&path, "数据", 3, 1).as_deref(), Some("SUM(B2:B3)"));
    }

    /// 2) 编辑后的单元格样式仍在（与编辑前的样式逐字段相等）
    #[test]
    fn keeps_cell_style_after_edit() {
        let dir = test_dir("style");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let before = style_at(&path, "数据", (2, 2)); // B2 = (col 2, row 2)（umya 坐标 1 起）
        assert!(before.font().map(|f| f.bold()).unwrap_or(false), "样本应是加粗样式");

        save(&path, None, &[edit("数据", 1, 1, "number", "42.5")], true).expect("保存应成功");
        let after = style_at(&path, "数据", (2, 2));
        assert_eq!(before, after, "改数字不应改动单元格样式");

        // 日期格式单元格写入序列号后同样保留日期格式
        save(&path, None, &[edit("数据", 4, 1, "date", "2024-01-01")], true).expect("保存应成功");
        assert_eq!(number_format_at(&path, "数据", (2, 5)), "yyyy-mm-dd");
    }

    /// 3) date 编辑：日期格式的格子写序列号，普通格子按字符串写
    #[test]
    fn date_edit_respects_cell_number_format() {
        let dir = test_dir("date");
        let path = dir.join("sample.xlsx");
        make_sample(&path);

        let edits = [
            edit("数据", 4, 1, "date", "2024-01-01"), // B5：日期格式 → 序列号 45292
            edit("数据", 5, 1, "date", "2024-01-01"), // B6：普通文本格 → 原样字符串
            edit("数据", 5, 3, "date", "1899-12-31"), // D6：1900-01-01 之前 → 原样字符串
        ];
        let result = save(&path, None, &edits, true).expect("保存应成功");
        assert_eq!(result.saved_cells, 3);

        // 有日期格式：写入的是数字序列号（calamine 会按日期类型返回）
        let serial = match read_cell(&path, "数据", 4, 1) {
            Data::DateTime(dt) => Some(dt.as_f64()),
            Data::Float(f) => Some(f),
            Data::Int(i) => Some(i as f64),
            _ => None,
        };
        assert_eq!(serial, Some(45292.0), "2024-01-01 应写成序列号 45292");
        assert_eq!(number_format_at(&path, "数据", (2, 5)), "yyyy-mm-dd", "日期格式应保留");

        // 没有日期格式：原文当字符串写入，用户不会看到 45292
        assert_eq!(
            read_cell(&path, "数据", 5, 1),
            Data::String("2024-01-01".into())
        );
        assert_eq!(
            read_cell(&path, "数据", 5, 3),
            Data::String("1899-12-31".into())
        );
    }

    /// 4) number 解析失败：返回 Err，且原文件字节不变
    #[test]
    fn invalid_number_leaves_file_untouched() {
        let dir = test_dir("bad-number");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let before = std::fs::read(&path).unwrap();

        let err = save(&path, None, &[edit("数据", 1, 1, "number", "十二")], true)
            .expect_err("非法数字应报错");
        assert!(err.contains("十二"), "错误信息应包含原值：{err}");
        assert!(err.contains("第 2 行第 2 列"), "错误信息应带坐标：{err}");
        assert_eq!(std::fs::read(&path).unwrap(), before, "失败时原文件不得改变");
        assert!(
            !dir.join("sample.xlsx.bak").exists(),
            "失败时不应留下备份文件"
        );
    }

    /// 5) 企业加密文件：全程内存加解密，输出仍是密文，解密后新值生效
    #[test]
    fn encrypted_workbook_stays_encrypted() {
        let dir = test_dir("encrypted");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let plain = std::fs::read(&path).unwrap();
        // 与真实加密文件一致的头：魔数 + 块大小 512 + 文件头长度 4096
        let mut header = vec![0u8; 4096];
        header[0..4].copy_from_slice(&[0xE0, 0xA8, 0x91, 0xE7]);
        header[8..12].copy_from_slice(&512u32.to_le_bytes());
        header[12..16].copy_from_slice(&4096u32.to_le_bytes());
        let wrapped = esafenet::encrypt_esafenet(&header, &plain);
        std::fs::write(&path, &wrapped).unwrap();

        let result = save(&path, None, &[edit("数据", 1, 0, "text", "香蕉")], true)
            .expect("加密工作簿应能保存");
        assert_eq!(result.saved_cells, 1);

        let out = std::fs::read(&path).unwrap();
        assert_eq!(&out[..4], &[0xE0, 0xA8, 0x91, 0xE7], "输出必须仍是加密格式");
        assert!(
            !out.windows(4).any(|w| w == b"PK\x03\x04"),
            "加密文件里不得出现明文 zip 内容"
        );
        // 备份的是加密后的原始字节，不是解密结果
        let backup = std::fs::read(dir.join("sample.xlsx.bak")).unwrap();
        assert_eq!(backup, wrapped, "备份应是原文件的密文字节");
        assert_eq!(&backup[..4], &[0xE0, 0xA8, 0x91, 0xE7]);

        let back = esafenet::decrypt_esafenet(&out).expect("落盘内容应能解密");
        let mut wb = open_workbook_auto_from_rs(Cursor::new(back)).expect("解密后应能解析");
        assert_eq!(cell_at(&mut wb, "数据", 1, 0), Data::String("香蕉".into()));
    }

    /// 6) 备份：默认生成且内容等于保存前的原始字节；backup=false 时不生成
    #[test]
    fn backup_written_by_default_and_skipped_when_disabled() {
        let dir = test_dir("backup");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let original = std::fs::read(&path).unwrap();
        let backup = dir.join("sample.xlsx.bak");

        let result = save(&path, None, &[edit("数据", 1, 0, "text", "香蕉")], true).unwrap();
        assert_eq!(
            result.backup_path.as_deref(),
            Some(backup.to_string_lossy().as_ref()),
            "应返回备份路径"
        );
        assert_eq!(
            std::fs::read(&backup).unwrap(),
            original,
            "备份内容应等于保存前的原始字节"
        );

        std::fs::remove_file(&backup).unwrap();
        let again = save(&path, None, &[edit("数据", 1, 0, "text", "梨")], false).unwrap();
        assert!(again.backup_path.is_none(), "未备份时 backupPath 应为 None");
        assert!(!backup.exists(), "backup=false 时不应生成备份");
        assert_eq!(read_cell(&path, "数据", 1, 0), Data::String("梨".into()));
    }

    /// 7) 保存后（含失败路径）目录里没有残留临时文件
    #[test]
    fn leaves_no_temp_files_behind() {
        let dir = test_dir("no-temp");
        let path = dir.join("sample.xlsx");
        make_sample(&path);

        save(&path, None, &[edit("数据", 1, 0, "text", "香蕉")], true).unwrap();
        // 失败路径同样不能留下临时文件
        let _ = save(&path, None, &[edit("数据", 1, 1, "number", "abc")], true);
        let _ = save(&path, None, &[edit("不存在", 0, 0, "text", "x")], true);

        let mut names: Vec<String> = std::fs::read_dir(&dir)
            .unwrap()
            .map(|e| e.unwrap().file_name().to_string_lossy().to_string())
            .collect();
        names.sort();
        assert_eq!(
            names,
            vec!["sample.xlsx".to_string(), "sample.xlsx.bak".to_string()],
            "目录里不应有残留文件：{names:?}"
        );
        assert!(!names.iter().any(|n| n.starts_with(TMP_PREFIX)));
    }

    /// 8) editable：只有 xlsx 为 true；非 xlsx 一律拒绝写回
    #[test]
    fn editable_only_for_xlsx() {
        assert!(is_editable(r"C:\a\报表.xlsx"));
        assert!(is_editable("REPORT.XLSX"));
        for other in ["a.xlsm", "a.xls", "a.xlsb", "a.ods", "a.csv", "noext", ""] {
            assert!(!is_editable(other), "{other} 不应可编辑");
        }
        let dir = test_dir("editable");
        let xlsx = dir.join("sample.xlsx");
        make_sample(&xlsx);
        // 同样内容的 .xlsm：整簿重写会把宏丢掉，必须拒绝
        let xlsm = dir.join("sample.xlsm");
        std::fs::copy(&xlsx, &xlsm).unwrap();
        let err = save(&xlsm, None, &[edit("数据", 1, 0, "text", "香蕉")], true)
            .expect_err("xlsm 不应允许写回");
        assert!(err.contains("只有 .xlsx"), "实际提示：{err}");
    }

    /// 另存为：目录必须存在、源文件不动、备份的是源文件；target 与 path 相同按普通保存
    #[test]
    fn save_as_behaviour() {
        let dir = test_dir("save-as");
        let path = dir.join("sample.xlsx");
        make_sample(&path);

        // 目录不存在 → 报错，且不留备份
        let missing = dir.join("nope").join("out.xlsx");
        let err = save(
            &path,
            Some(missing.to_str().unwrap()),
            &[edit("数据", 1, 0, "text", "香蕉")],
            true,
        )
        .expect_err("目录不存在应报错");
        assert!(err.contains("目录不存在"), "实际提示：{err}");
        assert!(!dir.join("sample.xlsx.bak").exists());

        // 正常另存为
        let target = dir.join("copy.xlsx");
        let source_before = std::fs::read(&path).unwrap();
        let result = save(
            &path,
            Some(target.to_str().unwrap()),
            &[edit("数据", 1, 0, "text", "香蕉")],
            true,
        )
        .expect("另存为应成功");
        assert_eq!(Path::new(&result.path), target.as_path());
        assert_eq!(
            read_cell(&path, "数据", 1, 0),
            Data::String("苹果".into()),
            "另存为不应改动源文件"
        );
        assert_eq!(
            read_cell(&target, "数据", 1, 0),
            Data::String("香蕉".into())
        );
        assert_eq!(
            std::fs::read(dir.join("sample.xlsx.bak")).unwrap(),
            source_before,
            "另存为备份的是源文件"
        );

        // target 与 path 相同 → 按普通保存处理
        let same_target = path.to_string_lossy().to_string();
        let same = save(
            &path,
            Some(same_target.as_str()),
            &[edit("数据", 1, 0, "text", "梨")],
            false,
        )
        .expect("target 与 path 相同应按普通保存处理");
        assert_eq!(Path::new(&same.path), path.as_path());
        assert_eq!(read_cell(&path, "数据", 1, 0), Data::String("梨".into()));

        // 非 xlsx 扩展名 → 拒绝，避免写出“名字是 xls、内容是 xlsx”的假文件
        let bad = dir.join("out.xls");
        let err = save(&path, Some(bad.to_str().unwrap()), &[], true)
            .expect_err("非 xlsx 目标应拒绝");
        assert!(err.contains("只能另存为 .xlsx"), "实际提示：{err}");
    }

    /// 字面值与「清空」编辑必须清掉该格原有公式
    #[test]
    fn literal_edit_clears_existing_formula() {
        let dir = test_dir("clear-formula");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        assert_eq!(
            read_formula(&path, "数据", 1, 2).as_deref(),
            Some("SUM(B2:B2)"),
            "样本里 C2 应有公式"
        );

        save(&path, None, &[edit("数据", 1, 2, "text", "备注")], true).unwrap();
        assert_eq!(read_formula(&path, "数据", 1, 2), None, "旧公式应被清除");
        assert_eq!(cell_at(&mut open(&path), "数据", 1, 2), Data::String("备注".into()));

        save(&path, None, &[edit("数据", 3, 1, "empty", "")], true).unwrap();
        assert_eq!(read_formula(&path, "数据", 3, 1), None, "清空也应清掉公式");
        assert_eq!(cell_at(&mut open(&path), "数据", 3, 1), Data::Empty);
    }

    /// 工作表名不存在 → 明确报错（含表名），文件不变
    #[test]
    fn unknown_sheet_reports_error_with_name() {
        let dir = test_dir("unknown-sheet");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let before = std::fs::read(&path).unwrap();
        let err = save(&path, None, &[edit("不存在", 0, 0, "text", "x")], true)
            .expect_err("表名不存在应报错");
        assert!(err.contains("不存在"), "错误信息应包含表名：{err}");
        assert_eq!(std::fs::read(&path).unwrap(), before);
    }

    /// 不支持的 kind / 非法坐标：报错而不是写坏文件
    #[test]
    fn rejects_unknown_kind_and_out_of_range_coordinates() {
        let dir = test_dir("bad-kind");
        let path = dir.join("sample.xlsx");
        make_sample(&path);
        let before = std::fs::read(&path).unwrap();

        let err = save(&path, None, &[edit("数据", 0, 0, "color", "#f00")], true)
            .expect_err("未知类型应报错");
        assert!(err.contains("color"), "错误信息应含类型：{err}");

        let err = save(&path, None, &[edit("数据", u32::MAX, 0, "text", "x")], true)
            .expect_err("越界坐标应报错");
        assert!(err.contains("越界"), "错误信息应说明越界：{err}");
        assert_eq!(std::fs::read(&path).unwrap(), before);
    }

    /// 日期正反换算：与 office.rs::format_serial 的锚点一致（互为反函数）
    #[test]
    fn serial_from_text_matches_known_anchors() {
        assert_eq!(serial_from_text("1900-01-01"), Some(1.0));
        assert_eq!(serial_from_text("1900-02-28"), Some(59.0));
        assert_eq!(serial_from_text("1900-03-01"), Some(61.0));
        assert_eq!(serial_from_text("1970-01-01"), Some(25569.0));
        assert_eq!(serial_from_text("2023-01-01"), Some(44927.0));
        assert_eq!(serial_from_text("2024-01-01"), Some(45292.0));
        assert_eq!(serial_from_text("2024/1/1"), Some(45292.0));
        assert_eq!(serial_from_text("2024-01-01 12:00"), Some(45292.5));
        assert_eq!(
            serial_from_text("2024/1/1 6:00:30"),
            Some(45292.0 + 6.0 / 24.0 + 30.0 / 86_400.0)
        );
        assert_eq!(serial_from_text("2024-01-01T00:00"), Some(45292.0));
        // 1900-01-01 之前不解析（当字符串写）
        assert_eq!(serial_from_text("1899-12-31"), None);
        assert_eq!(serial_from_text("1899-12-30"), None);
        // 不存在的日期 / 乱码 / 空值
        assert_eq!(serial_from_text("2024-02-31"), None);
        assert_eq!(serial_from_text("2024-13-01"), None);
        assert_eq!(serial_from_text("随便"), None);
        assert_eq!(serial_from_text(""), None);
        assert_eq!(serial_from_text("2024-01-01 25:00"), None);
    }

    /// 数字格式码 → 是否日期/时间格式
    #[test]
    fn detects_date_number_formats() {
        for code in [
            "yyyy-mm-dd",
            "m/d/yyyy",
            "dd/mm/yyyy",
            "h:mm:ss",
            "h:mm:ss;@",
            "[h]:mm:ss",
            "yyyy\"年\"m\"月\"",
            "mm-dd-yy",
        ] {
            assert!(format_code_is_date(code), "{code} 应识别为日期格式");
        }
        for code in [
            "",
            "@",
            "General",
            "0",
            "0.00",
            "#,##0.00",
            "#,##0.00;[Red]-#,##0.00",
            "0.00\" m\"",
            "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\)",
        ] {
            assert!(!format_code_is_date(code), "{code} 不应识别为日期格式");
        }
    }

    /// 数字 / 布尔文本解析
    #[test]
    fn parses_number_and_bool_text() {
        assert_eq!(parse_number(" 42.5 "), Some(42.5));
        assert_eq!(parse_number("-3"), Some(-3.0));
        assert_eq!(parse_number("1e3"), Some(1000.0));
        assert_eq!(parse_number("abc"), None);
        assert_eq!(parse_number(""), None);
        assert_eq!(parse_number("NaN"), None);
        assert_eq!(parse_number("inf"), None);

        assert_eq!(parse_bool("true"), Some(true));
        assert_eq!(parse_bool("TRUE"), Some(true));
        assert_eq!(parse_bool("1"), Some(true));
        assert_eq!(parse_bool("false"), Some(false));
        assert_eq!(parse_bool("FALSE"), Some(false));
        assert_eq!(parse_bool("0"), Some(false));
        assert_eq!(parse_bool("是"), None);
    }

    /// 真实样本（本机存在时才跑）：企业加密 xlsx 的「读 → 改 → 存 → 复读」完整往返。
    /// 全程只操作临时目录里的副本，绝不碰用户原文件；用于验证整条写回链路
    /// （内存解密 → umya 改写 → calamine 校验 → 重新加密 → .bak 备份 → 原子替换）。
    #[test]
    fn real_encrypted_workbook_roundtrip_if_present() {
        let real = r"Z:\D\mywork\01_project\007_topology_identification\06_data\voltage_daily_v3\2026-03-01.xlsx";
        if !Path::new(real).exists() {
            return;
        }
        let dir = test_dir("real-roundtrip");
        let copy = dir.join("copy.xlsx");
        std::fs::copy(real, &copy).expect("复制真实样本到临时目录");
        let before = std::fs::read(&copy).expect("读取副本");
        assert!(
            esafenet::is_esafenet_encrypted(&before),
            "真实样本应是企业加密文档"
        );

        // 解密后再读（模拟应用打开流程）
        let read = |path: &Path, sheet: &str, row: u32, col: u32| -> Data {
            let raw = std::fs::read(path).expect("读取文件");
            let plain = esafenet::decrypt_esafenet(&raw).unwrap_or(raw);
            let mut wb = open_workbook_auto_from_rs(Cursor::new(plain)).expect("应能解析工作簿");
            cell_at(&mut wb, sheet, row, col)
        };
        let sheet = {
            let raw = std::fs::read(&copy).unwrap();
            let plain = esafenet::decrypt_esafenet(&raw).unwrap_or(raw);
            let wb = open_workbook_auto_from_rs(Cursor::new(plain)).expect("应能解析工作簿");
            wb.sheet_names().first().cloned().expect("至少一个工作表")
        };
        let original = read(&copy, &sheet, 0, 0);
        assert!(
            matches!(original, Data::String(_)),
            "A1 应为文本，实际 {original:?}"
        );

        // 改两格：A1 文本、A2 数字
        let marker = "MasterEdit 往返测试";
        let saved = save(
            &copy,
            None,
            &[
                edit(&sheet, 0, 0, "text", marker),
                edit(&sheet, 1, 0, "number", "12345"),
            ],
            true,
        )
        .expect("真实加密工作簿应能保存");
        assert_eq!(saved.saved_cells, 2);

        // 加密壳仍在，备份是原始密文字节
        let after = std::fs::read(&copy).expect("读取保存后的文件");
        assert!(
            esafenet::is_esafenet_encrypted(&after),
            "保存后仍应是加密文档"
        );
        let backup = saved.backup_path.clone().expect("应生成 .bak 备份");
        assert_eq!(
            std::fs::read(&backup).expect("读取备份"),
            before,
            "备份应为原始密文字节"
        );

        // 复读：新值生效
        assert_eq!(read(&copy, &sheet, 0, 0), Data::String(marker.into()));
        let numeric = read(&copy, &sheet, 1, 0);
        let numeric_ok = match numeric {
            Data::Float(f) => (f - 12345.0).abs() < 1e-6,
            Data::Int(i) => i == 12345,
            _ => false,
        };
        assert!(numeric_ok, "A2 应为数字 12345，实际 {numeric:?}");

        // 工作表结构不变（真实样本 300+ 行）
        let rows_after = {
            let raw = std::fs::read(&copy).unwrap();
            let plain = esafenet::decrypt_esafenet(&raw).unwrap_or(raw);
            let mut wb = open_workbook_auto_from_rs(Cursor::new(plain)).unwrap();
            wb.worksheet_range(&sheet).unwrap().height()
        };
        assert!(rows_after > 300, "真实样本应有 300+ 行，实际 {rows_after}");

        // 不残留临时文件
        let leftovers: Vec<String> = std::fs::read_dir(&dir)
            .unwrap()
            .filter_map(|entry| entry.ok())
            .map(|entry| entry.file_name().to_string_lossy().to_string())
            .filter(|name| name.starts_with(".MasterEdit-tmp-"))
            .collect();
        assert!(leftovers.is_empty(), "不应残留临时文件: {leftovers:?}");
    }
}
