//! 文件读写相关命令：读取 / 写入 / 保存对话框 / 路径辅助。

use crate::commands::{esafenet, textcodec};
use serde::Serialize;
use std::path::{Path, PathBuf};
use std::time::UNIX_EPOCH;

/// 前端读取文件后拿到的数据。
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FilePayload {
    pub path: String,
    pub content: String,
    pub modified_at: u64,
    pub size: u64,
    /// 是否为 Esafenet 透明加密文档（内容已解密，保存时按原格式加密写回）
    pub encrypted: bool,
    /// 加密文档的 4096 字节文件头（base64），保存时用于重新加密
    pub encrypted_header: Option<String>,
    /// 文件编码（检测结果或指定值）
    pub encoding: String,
    /// 换行符：lf / crlf / cr
    pub eol: String,
}

/// 校验并规范化用户传入的路径，拒绝空路径与含 NUL 的路径。
pub fn validate_path(path: &str) -> Result<PathBuf, String> {
    if path.trim().is_empty() {
        return Err("路径为空".to_string());
    }
    if path.contains('\0') {
        return Err("非法路径".to_string());
    }
    Ok(PathBuf::from(path))
}

/// 取文件的最后修改时间（Unix 毫秒）。
pub fn modified_ms(meta: &std::fs::Metadata) -> u64 {
    meta.modified()
        .ok()
        .and_then(|t| t.duration_since(UNIX_EPOCH).ok())
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}

#[tauri::command]
pub async fn read_markdown_file(
    path: String,
    encoding: Option<String>,
) -> Result<FilePayload, String> {
    let p = validate_path(&path)?;
    let meta = std::fs::metadata(&p).map_err(|e| format!("无法读取文件信息: {e}"))?;
    if meta.is_dir() {
        return Err("目标是文件夹，无法作为文档打开".to_string());
    }
    let raw = std::fs::read(&p).map_err(|e| format!("读取文件失败: {e}"))?;
    // 企业透明加密文档：自动解密为明文供编辑预览
    let (bytes, encrypted, header) = match esafenet::decrypt_esafenet(&raw) {
        Some(decrypted) => (decrypted, true, Some(base64_encode(&raw[..4096]))),
        None => (raw, false, None),
    };
    // 编码：优先使用调用方指定值（用于「以其它编码重新解释」），否则自动检测
    let enc = encoding
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| textcodec::detect_encoding(&bytes).to_string());
    let content = textcodec::decode_with(&bytes, &enc);
    let eol = textcodec::detect_eol(&content).to_string();
    Ok(FilePayload {
        path: p.to_string_lossy().to_string(),
        content,
        modified_at: modified_ms(&meta),
        size: meta.len(),
        encrypted,
        encrypted_header: header,
        encoding: enc,
        eol,
    })
}

/// 写入文件。返回写入后的修改时间（毫秒）供前端刷新基线。
///
/// - 换行符：按 eol 统一转换（默认 lf）
/// - 编码：按 encoding 编码（默认 utf-8）
/// - 加密文档保持加密格式写回：
///   1. 前端传入打开时的 4096 字节文件头（base64）时，用该头重新加密；
///   2. 否则若目标文件本身是加密文档（如新建文档覆盖加密文件），沿用其文件头；
///   3. 都不是时按普通文本原样写入。
#[tauri::command]
pub async fn write_markdown_file(
    path: String,
    content: String,
    encrypted_header: Option<String>,
    encoding: Option<String>,
    eol: Option<String>,
) -> Result<u64, String> {
    let p = validate_path(&path)?;
    if let Some(parent) = p.parent() {
        if !parent.as_os_str().is_empty() && !parent.exists() {
            std::fs::create_dir_all(parent).map_err(|e| format!("创建目录失败: {e}"))?;
        }
    }

    let header = match encrypted_header.filter(|s| !s.is_empty()) {
        Some(b64) => Some(base64_decode(&b64)?),
        None => {
            // 目标已存在且为加密文档时沿用其头部
            std::fs::read(&p)
                .ok()
                .filter(|existing| esafenet::is_esafenet_encrypted(existing))
                .map(|existing| existing[..4096].to_vec())
        }
    };

    let enc = encoding.filter(|s| !s.is_empty()).unwrap_or_else(|| "utf-8".to_string());
    let normalized = textcodec::normalize_eol(&content, eol.as_deref().unwrap_or("lf"));
    let plain = textcodec::encode_with(&normalized, &enc)?;

    let bytes = match header {
        Some(head) => esafenet::encrypt_esafenet(&head, &plain),
        None => plain,
    };

    std::fs::write(&p, &bytes).map_err(|e| format!("写入文件失败: {e}"))?;
    let meta = std::fs::metadata(&p).map_err(|e| format!("读取文件信息失败: {e}"))?;
    Ok(modified_ms(&meta))
}

/// 前端读取二进制文件后拿到的数据（用于 PDF 等）。
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BinaryPayload {
    pub path: String,
    pub data_base64: String,
    pub modified_at: u64,
    pub size: u64,
    /// 是否为 Esafenet 透明加密文档（内容已解密，保存时按原格式加密写回）
    pub encrypted: bool,
    /// 加密文档的 4096 字节文件头（base64），保存时用于重新加密
    pub encrypted_header: Option<String>,
}

#[tauri::command]
pub async fn read_binary_file(path: String) -> Result<BinaryPayload, String> {
    let p = validate_path(&path)?;
    let meta = std::fs::metadata(&p).map_err(|e| format!("无法读取文件信息: {e}"))?;
    if meta.is_dir() {
        return Err("目标是文件夹，无法作为文档打开".to_string());
    }
    let raw = std::fs::read(&p).map_err(|e| format!("读取文件失败: {e}"))?;
    // 企业透明加密文档：自动解密为明文供查看/编辑
    let (bytes, encrypted, header) = match esafenet::decrypt_esafenet(&raw) {
        Some(decrypted) => (decrypted, true, Some(base64_encode(&raw[..4096]))),
        None => (raw, false, None),
    };
    Ok(BinaryPayload {
        path: p.to_string_lossy().to_string(),
        data_base64: base64_encode(&bytes),
        modified_at: modified_ms(&meta),
        size: meta.len(),
        encrypted,
        encrypted_header: header,
    })
}

/// 弹出「另存为」系统对话框。前端负责把返回值写入文件。
#[tauri::command]
pub async fn save_file_dialog(
    app: tauri::AppHandle,
    default_path: Option<String>,
    filter_all: Option<bool>,
    filter_pdf: Option<bool>,
) -> Result<Option<String>, String> {
    use tauri_plugin_dialog::DialogExt;

    let is_pdf = filter_pdf.unwrap_or(false)
        || default_path
            .as_deref()
            .map(|p| p.to_lowercase().ends_with(".pdf"))
            .unwrap_or(false);
    // 表格另存为：按默认文件名后缀识别，无需新增参数
    let is_sheet = default_path
        .as_deref()
        .map(|p| {
            let lower = p.to_lowercase();
            lower.ends_with(".xlsx")
                || lower.ends_with(".xlsm")
                || lower.ends_with(".xls")
                || lower.ends_with(".ods")
        })
        .unwrap_or(false);

    let mut builder = app.dialog().file().set_title("另存为");
    if is_pdf {
        builder = builder
            .add_filter("PDF 文件", &["pdf"])
            .add_filter("所有文件", &["*"]);
    } else if is_sheet {
        builder = builder
            .add_filter("Excel 工作簿", &["xlsx", "xlsm"])
            .add_filter("所有文件", &["*"]);
    } else if filter_all.unwrap_or(false) {
        builder = builder
            .add_filter("所有文件", &["*"])
            .add_filter("Markdown 文件", &["md", "markdown", "mdown"])
            .add_filter("文本文件", &["txt"]);
    } else {
        builder = builder
            .add_filter("Markdown 文件", &["md", "markdown", "mdown"])
            .add_filter("文本文件", &["txt"])
            .add_filter("所有文件", &["*"]);
    }

    if let Some(dp) = default_path.filter(|s| !s.is_empty()) {
        let path = PathBuf::from(&dp);
        if let Some(dir) = path.parent().and_then(|d| d.to_str()) {
            if !dir.is_empty() {
                builder = builder.set_directory(dir);
            }
        }
        if let Some(name) = path.file_name().and_then(|n| n.to_str()) {
            builder = builder.set_file_name(name);
        }
    }

    let (tx, rx) = std::sync::mpsc::channel();
    builder.save_file(move |file| {
        let _ = tx.send(file);
    });
    match rx.recv() {
        Ok(Some(file)) => {
            let p = file
                .into_path()
                .map_err(|e| format!("无效的保存路径: {e}"))?;
            Ok(Some(p.to_string_lossy().to_string()))
        }
        Ok(None) => Ok(None),
        Err(e) => Err(format!("保存对话框失败: {e}")),
    }
}

/// 判断路径是否存在（用于最近文件失效检测 / 图片解析）。
#[tauri::command]
pub async fn path_exists(path: String) -> Result<bool, String> {
    let p = validate_path(&path)?;
    Ok(p.exists())
}

/// 读取任意文件并返回 base64（导出 HTML 时内联图片使用）。
#[tauri::command]
pub async fn read_file_as_base64(path: String) -> Result<String, String> {
    let p = validate_path(&path)?;
    let bytes = std::fs::read(&p).map_err(|e| format!("读取失败: {e}"))?;
    Ok(base64_encode(&bytes))
}

const B64_TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/// 无依赖 base64 编码（标准字母表，带 padding）。
fn base64_encode(input: &[u8]) -> String {
    let mut out = String::with_capacity(input.len().div_ceil(3) * 4);
    for chunk in input.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = *chunk.get(1).unwrap_or(&0) as u32;
        let b2 = *chunk.get(2).unwrap_or(&0) as u32;
        let n = (b0 << 16) | (b1 << 8) | b2;
        out.push(B64_TABLE[(n >> 18) as usize & 63] as char);
        out.push(B64_TABLE[(n >> 12) as usize & 63] as char);
        out.push(if chunk.len() > 1 {
            B64_TABLE[(n >> 6) as usize & 63] as char
        } else {
            '='
        });
        out.push(if chunk.len() > 2 {
            B64_TABLE[n as usize & 63] as char
        } else {
            '='
        });
    }
    out
}

/// 解码标准 base64（接受换行/空白，忽略非法字符）。
fn base64_decode(input: &str) -> Result<Vec<u8>, String> {
    let cleaned: Vec<u8> = input
        .bytes()
        .filter(|b| !b.is_ascii_whitespace())
        .collect();
    if cleaned.len() % 4 != 0 {
        return Err("base64 长度非法".to_string());
    }
    let value_of = |b: u8| -> Option<u32> {
        match b {
            b'A'..=b'Z' => Some((b - b'A') as u32),
            b'a'..=b'z' => Some((b - b'a') as u32 + 26),
            b'0'..=b'9' => Some((b - b'0') as u32 + 52),
            b'+' => Some(62),
            b'/' => Some(63),
            _ => None,
        }
    };
    let mut out = Vec::with_capacity(cleaned.len() / 4 * 3);
    for chunk in cleaned.chunks(4) {
        let mut n: u32 = 0;
        let mut pad = 0;
        for (i, &b) in chunk.iter().enumerate() {
            if b == b'=' {
                pad += 1;
                n <<= 6;
                continue;
            }
            let v = value_of(b).ok_or_else(|| "base64 含非法字符".to_string())?;
            n = (n << 6) | v;
            let _ = i;
        }
        out.push((n >> 16) as u8);
        if pad < 2 {
            out.push((n >> 8) as u8);
        }
        if pad < 1 {
            out.push(n as u8);
        }
    }
    Ok(out)
}

/// 写入二进制文件（PDF 保存 / 导出 PNG / DOCX 使用），内容以 base64 传输。
#[tauri::command]
pub async fn write_binary_file(
    path: String,
    base64: String,
    encrypted_header: Option<String>,
) -> Result<u64, String> {
    let p = validate_path(&path)?;
    if let Some(parent) = p.parent() {
        if !parent.as_os_str().is_empty() && !parent.exists() {
            std::fs::create_dir_all(parent).map_err(|e| format!("创建目录失败: {e}"))?;
        }
    }
    let header = match encrypted_header.filter(|s| !s.is_empty()) {
        Some(b64) => Some(base64_decode(&b64)?),
        None => {
            // 目标已存在且为加密文档时沿用其头部
            std::fs::read(&p)
                .ok()
                .filter(|existing| esafenet::is_esafenet_encrypted(existing))
                .map(|existing| existing[..4096].to_vec())
        }
    };
    let plain = base64_decode(&base64)?;
    let bytes = match header {
        Some(head) => esafenet::encrypt_esafenet(&head, &plain),
        None => plain,
    };
    std::fs::write(&p, &bytes).map_err(|e| format!("写入文件失败: {e}"))?;
    let meta = std::fs::metadata(&p).map_err(|e| format!("读取文件信息失败: {e}"))?;
    Ok(modified_ms(&meta))
}

/// 供其它模块复用的路径父目录提取。
pub fn parent_dir(path: &str) -> Result<PathBuf, String> {
    let p = validate_path(path)?;
    p.parent()
        .filter(|d| !d.as_os_str().is_empty())
        .map(Path::to_path_buf)
        .ok_or_else(|| "无法确定文件所在目录".to_string())
}
