mod commands;

use commands::{file, image, recent, watch};
use std::sync::Mutex;
use tauri::{Emitter, Manager, WindowEvent};

/// 各窗口中是否存在未保存变更（由各窗口前端同步自己的状态）。
///
/// 必须**按窗口**记录：旧实现用一个全局 AtomicBool，打开第二个窗口时新窗口挂载即写入
/// `set_dirty(false)`，把主窗口的「有未保存内容」标记清掉，
/// 于是主窗口点关闭时不再拦截、不再弹确认框，内存中的编辑直接丢失。
pub struct DirtyWindows(pub Mutex<std::collections::HashSet<String>>);

/// 通过文件关联双击启动时，命令行传入的待打开文件。
pub struct StartupFile(pub Mutex<Option<String>>);

#[tauri::command]
fn set_dirty(
    window: tauri::Window,
    state: tauri::State<'_, DirtyWindows>,
    dirty: bool,
) {
    let label = window.label().to_string();
    if let Ok(mut set) = state.0.lock() {
        if dirty {
            set.insert(label);
        } else {
            set.remove(&label);
        }
    }
}

#[tauri::command]
fn get_startup_file(state: tauri::State<'_, StartupFile>) -> Option<String> {
    state.0.lock().ok().and_then(|g| g.clone())
}

/// 前端确认可以关闭后，真正销毁窗口。
#[tauri::command]
fn confirm_close(window: tauri::Window) -> Result<(), String> {
    window.destroy().map_err(|e| e.to_string())
}

fn percent_encode(s: &str) -> String {
    let mut out = String::new();
    for b in s.as_bytes() {
        match *b {
            b'a'..=b'z' | b'A'..=b'Z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => {
                out.push(*b as char);
            }
            _ => {
                use std::fmt::Write;
                let _ = write!(out, "%{:02X}", b);
            }
        }
    }
    out
}

/// 在独立的新窗口中打开文档
#[tauri::command]
async fn open_in_new_window(app: tauri::AppHandle, path: String) -> Result<(), String> {
    let title = std::path::Path::new(&path)
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "MasterEdit".to_string());
    let encoded = percent_encode(&path);
    let label = format!(
        "win-{}",
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis()
    );
    let url_str = format!("index.html?open={}", encoded);
    let _window = tauri::WebviewWindowBuilder::new(
        &app,
        label,
        tauri::WebviewUrl::App(url_str.into()),
    )
    .title(title)
    .inner_size(1000.0, 760.0)
    .build()
    .map_err(|e| format!("打开新窗口失败: {e}"))?;

    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_store::Builder::new().build())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(DirtyWindows(Mutex::new(std::collections::HashSet::new())))
        .setup(|app| {
            let handle = app.handle().clone();
            app.manage(recent::load(&handle));
            app.manage(watch::WatchState::default());
            // 改名兼容：把旧品牌 MasterMD 的存储文件迁移到新文件名（保留 PDF 批注等数据）
            commands::legacy::migrate_legacy_store(&handle);

            // 文件关联：Windows 会把文件路径作为首个非选项参数传入
            let startup = std::env::args()
                .skip(1)
                .find(|a| !a.starts_with('-') && std::path::Path::new(a).exists());
            app.manage(StartupFile(Mutex::new(startup)));

            // 关闭 WebView2 自带快捷键，避免与应用的 Ctrl+F / Ctrl+P / F3 冲突
            #[cfg(windows)]
            commands::pdf::disable_browser_accelerators(&handle);
            // 清理便携版自更新遗留的旧程序备份
            #[cfg(windows)]
            commands::update::cleanup_old_binary();
            let _ = &handle;
            Ok(())
        })
        .on_window_event(|window, event| match event {
            WindowEvent::CloseRequested { api, .. } => {
                // 只看当前窗口自己的未保存标记
                let dirty = window
                    .state::<DirtyWindows>()
                    .0
                    .lock()
                    .map(|set| set.contains(window.label()))
                    .unwrap_or(false);
                if dirty {
                    // 有未保存内容：拦截关闭，交由前端弹窗确认
                    api.prevent_close();
                    let _ = window.emit("close-requested", ());
                }
            }
            WindowEvent::Destroyed => {
                // 窗口已销毁，清掉它的标记，避免残留
                if let Ok(mut set) = window.state::<DirtyWindows>().0.lock() {
                    set.remove(window.label());
                }
            }
            _ => {}
        })
        .invoke_handler(tauri::generate_handler![
            set_dirty,
            get_startup_file,
            confirm_close,
            open_in_new_window,
            file::read_markdown_file,
            file::read_binary_file,
            file::write_markdown_file,
            file::write_binary_file,
            file::save_file_dialog,
            file::path_exists,
            file::read_file_as_base64,
            commands::fs::list_directory,
            commands::fs::parent_dir_of,
            commands::fs::search_workspace_symbols,
            recent::get_recent_files,
            recent::add_recent_file,
            recent::remove_recent_file,
            recent::clear_recent_files,
            image::save_pasted_image,
            watch::watch_file,
            watch::unwatch_file,
            #[cfg(windows)]
            commands::pdf::print_to_pdf,
            #[cfg(windows)]
            commands::update::check_update,
            #[cfg(windows)]
            commands::update::download_update,
            #[cfg(windows)]
            commands::update::run_installer,
            #[cfg(windows)]
            commands::update::apply_installer_update,
            #[cfg(windows)]
            commands::update::apply_update,
            #[cfg(windows)]
            commands::update::quit_app,
            #[cfg(windows)]
            commands::update::current_exe_path,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
