//! 粘贴图片落盘：在文档同目录创建 assets/ 并写入 PNG/JPG 等。

use crate::commands::file::{parent_dir, validate_path};
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

/// 规范化扩展名，仅允许常见图片格式。
fn normalize_ext(ext: &str) -> String {
    let e = ext
        .trim_start_matches('.')
        .to_ascii_lowercase()
        .replace(|c: char| !c.is_ascii_alphanumeric(), "");
    match e.as_str() {
        "png" | "jpg" | "jpeg" | "gif" | "webp" | "bmp" | "svg" => e,
        _ => "png".to_string(),
    }
}

/// 将 Unix 天数拆成 UTC 的 (年, 月, 日)。office.rs 解析 Excel 序列日期时复用。
pub(crate) fn civil_from_days(z: i64) -> (i64, u32, u32) {
    // Howard Hinnant 的 civil_from_days 算法
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

fn timestamp() -> String {
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);
    let (y, mo, d, h, mi, s) = {
        let days = secs.div_euclid(86_400);
        let rem = secs.rem_euclid(86_400);
        let (y, mo, d) = civil_from_days(days);
        (y, mo, d, rem / 3600, (rem % 3600) / 60, rem % 60)
    };
    format!("{y:04}{mo:02}{d:02}-{h:02}{mi:02}{s:02}")
}

/// 保存剪贴板图片，返回可直接插入 Markdown 的相对路径（如 assets/image-20260101-120000.png）。
#[tauri::command]
pub async fn save_pasted_image(
    doc_path: String,
    data: Vec<u8>,
    ext: String,
) -> Result<String, String> {
    if data.is_empty() {
        return Err("图片数据为空".to_string());
    }
    validate_path(&doc_path)?;
    let dir = parent_dir(&doc_path)?;
    let assets = dir.join("assets");    std::fs::create_dir_all(&assets).map_err(|e| format!("创建 assets 目录失败: {e}"))?;

    let ext = normalize_ext(&ext);
    let base = format!("image-{}", timestamp());
    let mut name = format!("{base}.{ext}");
    let mut idx = 1;
    while assets.join(&name).exists() {
        name = format!("{base}-{idx}.{ext}");
        idx += 1;
        if idx > 999 {
            return Err("同名图片过多".to_string());
        }
    }
    let target: PathBuf = assets.join(&name);
    std::fs::write(&target, &data).map_err(|e| format!("保存图片失败: {e}"))?;
    Ok(format!("assets/{name}"))
}
