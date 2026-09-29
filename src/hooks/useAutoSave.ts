import { useEffect } from "react";
import { useAppStore } from "../stores/appStore";
import { useSettingsStore } from "../stores/settingsStore";
import { saveDoc } from "../utils/fileActions";

/**
 * 自动保存：仅在文档已有磁盘路径、处于脏状态且非只读时触发。
 * 内容变化或脏状态变化会重新计时。
 */
export function useAutoSave(): void {
  const autoSave = useSettingsStore((s) => s.autoSave);
  const interval = useSettingsStore((s) => s.autoSaveInterval);
  const doc = useAppStore((s) => s.docs.find((d) => d.id === s.activeId) ?? null);

  const docId = doc?.id ?? null;
  const filePath = doc?.filePath ?? null;
  const dirty = doc?.isDirty ?? false;
  const readOnly = doc?.readOnly ?? false;
  const contentLength = doc?.content.length ?? 0;
  /**
   * 表格（xlsx）写回是「整簿重写」，必须由用户显式 Ctrl+S 触发：
   * 自动保存会在用户没确认的情况下重写工作簿，因此这里跳过表格文档。
   */
  const isSheet = doc?.docType === "spreadsheet";

  useEffect(() => {
    if (!autoSave || !docId || !filePath || !dirty || readOnly || isSheet) return;
    const timer = window.setTimeout(() => {
      const current = useAppStore.getState().docs.find((d) => d.id === docId);
      if (current?.isDirty && current.filePath && !current.readOnly && current.docType !== "spreadsheet") {
        void saveDoc(current.id);
      }
    }, Math.max(5, interval) * 1000);
    return () => window.clearTimeout(timer);
  }, [autoSave, interval, docId, filePath, dirty, readOnly, contentLength, isSheet]);
}
