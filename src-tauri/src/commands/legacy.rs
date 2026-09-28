//! 旧品牌（MasterMD）遗留数据的兼容迁移。
//!
//! 应用改名为 MasterEdit 后，插件存储文件名也从 `mastermd-store.json` 改为
//! `masteredit-store.json`。该文件里不仅有设置与会话，还保存着**全部 PDF 高亮与批注**，
//! 直接改名会让老用户丢失数据，因此启动时做一次「存在即复制」的迁移：
//! 新文件不存在且旧文件存在时才复制，旧文件一律保留（回滚到旧版本仍可读到数据）。
//!
//! Windows 上 `app_data_dir` 与 `app_config_dir` 指向同一目录
//! （`%APPDATA%\<identifier>`），插件存储文件就在其中。

use std::path::Path;
use tauri::{AppHandle, Manager};

/// 改名前的插件存储文件名
pub const LEGACY_STORE_FILE: &str = "mastermd-store.json";
/// 当前的插件存储文件名（必须与前端 `src/utils/persist.ts` 中的 `STORE_FILE` 一致）
pub const CURRENT_STORE_FILE: &str = "masteredit-store.json";

/// 在指定目录内执行迁移，返回是否真的复制了文件。
/// 独立成纯函数，便于单元测试覆盖「不该迁移」的分支。
pub fn migrate_store_in(dir: &Path) -> std::io::Result<bool> {
    let current = dir.join(CURRENT_STORE_FILE);
    let legacy = dir.join(LEGACY_STORE_FILE);
    // 新文件已存在（含空文件）说明已经迁移过或本来就是新装：绝不复盖用户当前数据
    if current.exists() || !legacy.exists() {
        return Ok(false);
    }
    std::fs::copy(&legacy, &current)?;
    Ok(true)
}

/// 应用启动时调用：把旧存储文件迁移到新文件名。
pub fn migrate_legacy_store(app: &AppHandle) {
    let Ok(dir) = app.path().app_data_dir() else {
        return;
    };
    match migrate_store_in(&dir) {
        Ok(true) => eprintln!("[migrate] 已将 {LEGACY_STORE_FILE} 迁移为 {CURRENT_STORE_FILE}"),
        Ok(false) => {}
        Err(error) => eprintln!("[migrate] 存储文件迁移失败: {error}"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn temp_dir(name: &str) -> std::path::PathBuf {
        let dir = std::env::temp_dir().join(format!("masteredit-migrate-{name}"));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).expect("创建临时测试目录");
        dir
    }

    #[test]
    fn copies_legacy_store_when_current_missing() {
        let dir = temp_dir("copy");
        std::fs::write(dir.join(LEGACY_STORE_FILE), r#"{"settings":{"fontSize":16}}"#).unwrap();

        assert!(migrate_store_in(&dir).expect("迁移应成功"));
        let migrated = std::fs::read_to_string(dir.join(CURRENT_STORE_FILE)).unwrap();
        assert!(migrated.contains("fontSize"));
        // 旧文件保留，便于回滚旧版本
        assert!(dir.join(LEGACY_STORE_FILE).exists());
    }

    #[test]
    fn never_overwrites_existing_current_store() {
        let dir = temp_dir("keep");
        std::fs::write(dir.join(LEGACY_STORE_FILE), "legacy").unwrap();
        std::fs::write(dir.join(CURRENT_STORE_FILE), "current").unwrap();

        assert!(!migrate_store_in(&dir).expect("应跳过迁移"));
        assert_eq!(
            std::fs::read_to_string(dir.join(CURRENT_STORE_FILE)).unwrap(),
            "current"
        );
    }

    #[test]
    fn noop_without_legacy_file() {
        let dir = temp_dir("noop");
        assert!(!migrate_store_in(&dir).expect("无旧文件时应跳过"));
        assert!(!dir.join(CURRENT_STORE_FILE).exists());
    }
}
