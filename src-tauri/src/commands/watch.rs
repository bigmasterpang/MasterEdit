//! 文件外部变更监听：监听文件所在目录，过滤出目标文件的事件并通过事件通道通知前端。

use notify::{EventKind, RecommendedWatcher, RecursiveMode, Watcher};
use serde::Serialize;
use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::Mutex;
use tauri::{AppHandle, Emitter, State};

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FileChangedPayload {
    pub path: String,
    pub exists: bool,
    pub modified_at: u64,
    pub size: u64,
}

/// 常驻状态：路径 -> watcher 实例（必须持有，否则监听会被释放）。
#[derive(Default)]
pub struct WatchState(pub Mutex<HashMap<String, RecommendedWatcher>>);

/// 去掉 Windows 扩展长度路径前缀（\\?\ / \\?\UNC\），保证与 notify 上报的路径一致
fn normalize(p: &std::path::Path) -> PathBuf {
    let s = p.to_string_lossy();
    if let Some(rest) = s.strip_prefix(r"\\?\UNC\") {
        return PathBuf::from(format!(r"\\{rest}"));
    }
    if let Some(rest) = s.strip_prefix(r"\\?\") {
        return PathBuf::from(rest);
    }
    p.to_path_buf()
}

fn key_of(path: &std::path::Path) -> String {
    normalize(path).to_string_lossy().to_lowercase()
}

#[tauri::command]
pub async fn watch_file(
    app: AppHandle,
    state: State<'_, WatchState>,
    path: String,
) -> Result<(), String> {
    let target = std::fs::canonicalize(&path)
        .map(|p| normalize(&p))
        .unwrap_or_else(|_| PathBuf::from(&path));
    let key = key_of(&target);
    {
        let map = state.0.lock().map_err(|e| e.to_string())?;
        if map.contains_key(&key) {
            return Ok(());
        }
    }
    // 监听父目录而非文件本身：编辑器常用的「原子替换保存」会导致文件句柄失效。
    let parent = target
        .parent()
        .map(|p| p.to_path_buf())
        .ok_or_else(|| "无法确定文件所在目录".to_string())?;

    let watched = target.clone();
    let emit_path = target.to_string_lossy().to_string();
    let mut watcher = notify::recommended_watcher(
        move |res: notify::Result<notify::Event>| {
            let Ok(event) = res else { return };
            if !matches!(
                event.kind,
                EventKind::Modify(_) | EventKind::Create(_) | EventKind::Remove(_)
            ) {
                return;
            }
            let watched_key = key_of(&watched);
            if !event.paths.iter().any(|p| key_of(p) == watched_key) {
                return;
            }
            let meta = std::fs::metadata(&watched).ok();
            let payload = FileChangedPayload {
                path: emit_path.clone(),
                exists: meta.is_some(),
                modified_at: meta
                    .as_ref()
                    .map(|m| crate::commands::file::modified_ms(m))
                    .unwrap_or(0),
                size: meta.as_ref().map(|m| m.len()).unwrap_or(0),
            };
            let _ = app.emit("file-changed", payload);
        },
    )
    .map_err(|e| format!("创建监听失败: {e}"))?;

    watcher
        .watch(&parent, RecursiveMode::NonRecursive)
        .map_err(|e| format!("监听目录失败: {e}"))?;

    state
        .0
        .lock()
        .map_err(|e| e.to_string())?
        .insert(key, watcher);
    Ok(())
}

#[tauri::command]
pub async fn unwatch_file(state: State<'_, WatchState>, path: String) -> Result<(), String> {
    let target = std::fs::canonicalize(&path)
        .map(|p| normalize(&p))
        .unwrap_or_else(|_| PathBuf::from(&path));
    let key = key_of(&target);
    let mut map = state.0.lock().map_err(|e| e.to_string())?;
    if map.remove(&key).is_some() {
        return Ok(());
    }
    // 文件已被删除或改名时 canonicalize 会失败，key 与注册时不一致；
    // 用原始路径再试一次，避免监听器永久残留在表里
    map.remove(&key_of(std::path::Path::new(&path)));
    Ok(())
}
