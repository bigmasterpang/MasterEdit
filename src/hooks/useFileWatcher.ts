import { useEffect } from "react";
import { listen } from "@tauri-apps/api/event";
import type { FileChangedPayload } from "../types";
import { useAppStore } from "../stores/appStore";
import { askConfirm, askConflict, showMessage } from "../stores/dialogStore";
import { fileName, samePath } from "../utils/filePath";
import {
  isRecentSelfWrite,
  reloadDocFromDisk,
  saveDocAs,
} from "../utils/fileActions";

/** 已被用户选择“仅查看当前内容”而忽略后续自动弹窗的文档 ID 集合 */
const ignoredReloadDocIds = new Set<string>();
/** 正在等待用户操作弹窗的文档 ID 集合（防止高频写入堆叠多个弹窗） */
const pendingPromptDocIds = new Set<string>();

/** 清除特定文档的忽略外部更新标记（在主动保存或手动重载时调用） */
export function clearIgnoredReload(docId: string): void {
  ignoredReloadDocIds.delete(docId);
  useAppStore.getState().patchDoc(docId, { ignoreExternalReload: false });
}

/**
 * 文件外部变更监听：
 * - 用户若选择“仅查看当前内容”，则不再对后续新更新重复弹窗
 * - 本地无未保存变更 -> 弹窗提示是否加载最新内容
 * - 本地有未保存变更 -> 弹窗提示冲突处理
 */
export function useFileWatcher(): void {
  useEffect(() => {
    let disposed = false;
    let unlisten: (() => void) | null = null;

    void listen<FileChangedPayload>("file-changed", async (event) => {
      const payload = event.payload;
      // 自己写入触发的事件直接忽略
      if (isRecentSelfWrite()) return;

      const app = useAppStore.getState();
      const doc = app.docs.find((d) => samePath(d.filePath, payload.path));
      if (!doc) return;

      if (!payload.exists) {
        await showMessage(
          "文件已被删除或移动",
          `「${fileName(payload.path)}」已不存在，请使用「另存为」保存到新的位置。`,
        );
        return;
      }
      if (payload.modifiedAt !== 0 && payload.modifiedAt === doc.modifiedAt) return;

      // 如果用户已选择“仅查看当前内容”，随着新的更新进来，不再重复弹窗
      if (doc.ignoreExternalReload || ignoredReloadDocIds.has(doc.id)) {
        useAppStore.getState().patchDoc(doc.id, {
          modifiedAt: payload.modifiedAt,
          size: payload.size,
        });
        return;
      }

      // 如果该文档已经有弹窗正在等待用户响应，避免重复堆叠弹窗
      if (pendingPromptDocIds.has(doc.id)) return;

      if (!doc.isDirty) {
        pendingPromptDocIds.add(doc.id);
        try {
          const reload = await askConfirm({
            title: "文件外部更新",
            message: `文档「${fileName(payload.path)}」在外部被修改，是否加载最新内容？`,
            confirmText: "重新加载",
            cancelText: "仅查看当前内容",
          });
          if (reload) {
            await reloadDocFromDisk(doc.id);
          } else {
            // 用户选择仅查看当前内容：随着新的更新进来，不再重复弹窗
            ignoredReloadDocIds.add(doc.id);
            useAppStore.getState().patchDoc(doc.id, {
              ignoreExternalReload: true,
              modifiedAt: payload.modifiedAt,
              size: payload.size,
            });
          }
        } finally {
          pendingPromptDocIds.delete(doc.id);
        }
        return;
      }

      pendingPromptDocIds.add(doc.id);
      try {
        const choice = await askConflict(fileName(payload.path));
        if (choice === "external") {
          await reloadDocFromDisk(doc.id);
        } else if (choice === "saveas") {
          await saveDocAs(doc.id);
        } else {
          // 保留本地：更新基线并标记忽略，后续外部改动不再重复弹窗
          ignoredReloadDocIds.add(doc.id);
          useAppStore.getState().patchDoc(doc.id, {
            ignoreExternalReload: true,
            modifiedAt: payload.modifiedAt,
            size: payload.size,
          });
        }
      } finally {
        pendingPromptDocIds.delete(doc.id);
      }
    }).then((fn) => {
      if (disposed) fn();
      else unlisten = fn;
    });

    return () => {
      disposed = true;
      unlisten?.();
    };
  }, []);
}
