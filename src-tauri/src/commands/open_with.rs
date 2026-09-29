//! 调用本地其它应用打开当前文档。
//!
//! 两个入口：
//! - `mode = "default"`：用系统默认关联程序打开（等同于在资源管理器里双击）；
//! - `mode = "choose"`：弹出 Windows 的「打开方式」选择框，让用户临时挑一个程序。
//!
//! 为什么用 ShellExecute 而不是自己拼命令行：
//! 1. `openas` 动词是系统提供的"打开方式"对话框，比自己枚举注册表可靠；
//! 2. 传的是**原始文件路径**，企业加密文档由系统驱动在目标程序里透明解密 ——
//!    我们**绝不**为了给外部程序看而写明文临时文件（这是项目的硬约定）。
//!
//! 注意：外部程序打开的是**磁盘上的版本**，编辑器里未保存的改动不在其中，
//! 前端调用前会先给出提示（有未保存改动时）。

use std::path::Path;

#[cfg(windows)]
use windows::core::PCWSTR;
#[cfg(windows)]
use windows::Win32::UI::Shell::ShellExecuteW;
#[cfg(windows)]
use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

/// 把路径转成 Windows API 需要的宽字符（以 NUL 结尾）
#[cfg(windows)]
fn wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(std::iter::once(0)).collect()
}

/// 用本地其它应用打开文件
///
/// `mode` 取值：`"default"`（系统默认程序）、`"choose"`（弹出「打开方式」选择框）。
#[tauri::command]
pub fn open_with_app(path: String, mode: Option<String>) -> Result<(), String> {
    let target = Path::new(&path);
    if !target.exists() {
        return Err(format!("文件不存在或已被移动：{path}"));
    }

    let verb = match mode.as_deref().unwrap_or("default") {
        "choose" => "openas",
        _ => "open",
    };

    #[cfg(windows)]
    {
        let verb_w = wide(verb);
        let path_w = wide(&path);
        // 返回值为 HINSTANCE：> 32 表示成功（这是 ShellExecute 的历史约定）
        let result = unsafe {
            ShellExecuteW(
                None,
                PCWSTR(verb_w.as_ptr()),
                PCWSTR(path_w.as_ptr()),
                PCWSTR::null(),
                PCWSTR::null(),
                SW_SHOWNORMAL,
            )
        };
        let code = result.0 as isize;
        if code <= 32 {
            return Err(match code {
                2 => "找不到关联的程序，请在「打开方式」里选择其它应用".to_string(),
                5 => "系统拒绝访问该文件（可能被安全软件拦截）".to_string(),
                _ => format!("调用外部程序失败（错误码 {code}）"),
            });
        }
        Ok(())
    }

    #[cfg(not(windows))]
    {
        let _ = verb;
        Err("当前平台暂不支持调用外部应用打开".to_string())
    }
}

#[cfg(all(test, windows))]
mod tests {
    use super::*;

    #[test]
    fn missing_file_reports_clear_error() {
        let err = open_with_app("C:/definitely/not/here.docx".into(), None).unwrap_err();
        assert!(err.contains("文件不存在"), "错误信息应说明文件不存在：{err}");
    }

    #[test]
    fn choose_mode_also_checks_file_first() {
        // 「打开方式」模式同样先校验文件存在，避免弹出对话框后才发现文件不在
        let err = open_with_app("C:/definitely/not/here.xlsx".into(), Some("choose".into())).unwrap_err();
        assert!(err.contains("文件不存在"), "错误信息应说明文件不存在：{err}");
    }
}
