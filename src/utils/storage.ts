/**
 * localStorage 兼容读取。
 *
 * 应用曾用名 MasterMD，早期版本把状态写在 `mastermd.*` 键下。改名后统一使用
 * `masteredit.*`，但直接换键等于清空用户数据（存档、江湖进度、资源管理器状态等），
 * 因此读取时回退旧键并顺手写入新键。旧键保留不删：用户回滚到旧版本仍能读到存档。
 *
 * localStorage 在隐私模式等场景可能抛异常，这里统一静默降级。
 */

/** 读取键值：优先新键，缺失时回退旧品牌键并把值迁移到新键 */
export function readMigratedItem(key: string, legacyKey: string): string | null {
  try {
    const value = localStorage.getItem(key);
    if (value !== null) return value;
    const legacy = localStorage.getItem(legacyKey);
    if (legacy === null) return null;
    localStorage.setItem(key, legacy);
    return legacy;
  } catch {
    return null;
  }
}

/** 写入键值（失败静默忽略） */
export function writeItem(key: string, value: string): void {
  try {
    localStorage.setItem(key, value);
  } catch {
    /* 忽略写入失败 */
  }
}

/** 删除键值：新旧键一起删，避免旧数据复活 */
export function removeItem(key: string, legacyKey: string): void {
  try {
    localStorage.removeItem(key);
    localStorage.removeItem(legacyKey);
  } catch {
    /* 忽略删除失败 */
  }
}
