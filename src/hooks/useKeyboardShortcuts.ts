import { useEffect } from "react";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { getActiveDoc, useAppStore } from "../stores/appStore";
import { getDocTitle, isMarkdownDoc } from "../utils/filePath";
import { useDialogStore } from "../stores/dialogStore";
import { useSearchStore } from "../stores/searchStore";
import { useSettingsStore } from "../stores/settingsStore";
import {
  deleteLines,
  duplicateLines,
  insertCallout,
  insertCodeBlock,
  insertHorizontalRule,
  insertTable,
  insertToc,
  moveLines,
  setHeading,
  shiftHeading,
  toggleBold,
  toggleHighlight,
  toggleInlineCode,
  toggleItalic,
  toggleList,
  toggleQuote,
  toggleStrikethrough,
  toggleSubscript,
  toggleSuperscript,
  toggleUnderline,
} from "../utils/editorCommands";
import {
  closeDocWithConfirm,
  newDocument,
  openFileDialog,
  saveActive,
  saveActiveAs,
} from "../utils/fileActions";
import {
  goToSymbolDefinition,
  peekSymbolDefinition,
  openSymbolReferences,
  navigateHistoryBack,
  navigateHistoryForward,
  useCodeNavStore,
} from "../utils/codeNavigation";
import { matchesShortcut } from "../utils/shortcuts";
import { APP_NAME } from "../utils/constants";
import type { ShortcutId } from "../types";

const BINDABLE: ShortcutId[] = [
  "open",
  "save",
  "saveAs",
  "newDoc",
  "viewMode",
  "search",
  "replace",
  "closeTab",
  "settings",
];

/**
 * 全局快捷键。
 * 可自定义项读取 settings.shortcuts；其余为编辑器固定快捷键（源码/分屏模式生效）。
 */
export function useKeyboardShortcuts(): void {
  useEffect(() => {
    const runAction = (id: ShortcutId) => {
      switch (id) {
        case "open":
          void openFileDialog();
          break;
        case "save":
          void saveActive();
          break;
        case "saveAs":
          void saveActiveAs();
          break;
        case "newDoc":
          void newDocument();
          break;
        case "viewMode": {
          // 非 Markdown 文档固定源码模式，不参与视图切换
          const active = getActiveDoc();
          if (isMarkdownDoc(active)) {
            useAppStore.getState().cycleViewMode();
          }
          break;
        }
        case "search":
          useSearchStore.getState().open(false);
          break;
        case "replace":
          useSearchStore.getState().open(true);
          break;
        case "closeTab": {
          const activeId = useAppStore.getState().activeId;
          if (activeId) void closeDocWithConfirm(activeId);
          break;
        }
        case "settings":
          useDialogStore.getState().setSettingsVisible(true);
          break;
      }
    };

    /** 编辑器类快捷键在可编辑状态下生效 */
    const editorReady = () => {
      const state = useAppStore.getState();
      const doc = getActiveDoc();
      return state.viewMode !== "preview" && Boolean(doc) && !doc?.readOnly;
    };

    /** Markdown 专属快捷键仅在 Markdown 文档下生效 */
    const markdownReady = () => {
      const state = useAppStore.getState();
      const doc = getActiveDoc();
      return (
        state.viewMode !== "preview" &&
        Boolean(doc) &&
        !doc?.readOnly &&
        isMarkdownDoc(doc)
      );
    };

    const handler = (event: KeyboardEvent) => {
      try {
        handleKey(event);
        // 已在处理函数中 preventDefault 的按键，阻止继续传播给 CodeMirror
        // 等内部处理器（否则 Alt+↑ 会被编辑器同时执行"上移行"）
        if (event.defaultPrevented) event.stopPropagation();
      } catch (error) {
        // 单个快捷键异常不应影响其它功能
        console.error("[MasterEdit] 快捷键处理失败", event.key, error);
      }
    };

    const handleKey = (event: KeyboardEvent) => {
      const mod = event.ctrlKey || event.metaKey;
      const key = event.key.toLowerCase();
      const code = event.code;
      const target = event.target as HTMLElement | null;

      // 快捷键录制面板打开时，除 Esc/F1 外一律不派发全局快捷键：
      // 面板里的录制输入框无法阻止窗口捕获阶段的监听器，
      // 否则「重新绑定 Ctrl+W」会在录制的同时真的把当前标签页关掉
      if (useDialogStore.getState().shortcutsVisible && key !== "escape" && event.key !== "F1") {
        return;
      }

      // 注意：CodeMirror 编辑区也是 contenteditable，不能当作"输入框"排除，
      // 否则在编辑器中按 Ctrl+B/I/K 会被直接吞掉。
      const inCodeMirror = Boolean(target?.closest?.(".cm-editor"));
      const inInput =
        !inCodeMirror &&
        (target?.tagName === "INPUT" ||
          target?.tagName === "TEXTAREA" ||
          target?.isContentEditable === true);

      if (!mod) {
        if (event.key === "F1") {
          event.preventDefault();
          useDialogStore.getState().setShortcutsVisible(true);
          return;
        }
        if (key === "escape") {
          const dialogs = useDialogStore.getState();
          if (
            dialogs.unsaved.open ||
            dialogs.conflict.open ||
            dialogs.confirm.open ||
            dialogs.message.open
          ) {
            return;
          }
          if (dialogs.settingsVisible) {
            dialogs.setSettingsVisible(false);
            return;
          }
          if (dialogs.shortcutsVisible) {
            dialogs.setShortcutsVisible(false);
            return;
          }
          if (useCodeNavStore.getState().peekState?.visible) {
            useCodeNavStore.getState().closePeek();
            return;
          }
          if (useSearchStore.getState().visible) {
            useSearchStore.getState().close();
          }
          return;
        }
        // F12：转到定义；Alt+F12：速览定义；Shift+F12：查找所有引用
        if (event.key === "F12") {
          const activeId = useAppStore.getState().activeId;
          if (!activeId) return;
          event.preventDefault();
          if (event.altKey) {
            void peekSymbolDefinition(activeId);
          } else if (event.shiftKey) {
            void openSymbolReferences(activeId);
          } else {
            void goToSymbolDefinition(activeId);
          }
          return;
        }
        // F3 / Shift+F3 查找下一个 / 上一个（Notepad++ 习惯）
        if (event.key === "F3") {
          event.preventDefault();
          const search = useSearchStore.getState();
          if (!search.visible) search.open(false);
          search.step(event.shiftKey ? -1 : 1);
          return;
        }
        // Alt+← / Alt+→ 代码导航后退 / 前进
        if (event.altKey && !event.shiftKey && (event.key === "ArrowLeft" || event.key === "ArrowRight")) {
          event.preventDefault();
          if (event.key === "ArrowLeft") {
            void navigateHistoryBack();
          } else {
            void navigateHistoryForward();
          }
          return;
        }
        // 聚焦左右栏：Alt+1 / Alt+2
        if (event.altKey && !event.shiftKey && (key === "1" || key === "2")) {
          event.preventDefault();
          useAppStore.getState().setActivePane(key === "1" ? 0 : 1);
          return;
        }
        // Alt+↑/↓ 上下移动行；Shift+Alt+↑/↓ 复制行
        if (event.altKey && (event.key === "ArrowUp" || event.key === "ArrowDown")) {
          if (!editorReady()) return;
          event.preventDefault();
          if (event.shiftKey) duplicateLines();
          else moveLines(event.key === "ArrowUp" ? -1 : 1);
          return;
        }
        return;
      }

      // 输入框内不劫持格式快捷键（编辑器除外）
      if (inInput && ["b", "i", "k", "`"].includes(key)) return;

      const shortcuts = useSettingsStore.getState().shortcuts;
      const matched = BINDABLE.find((id) => matchesShortcut(event, shortcuts[id] ?? ""));
      if (matched) {
        event.preventDefault();
        runAction(matched);
        return;
      }

      // ---------------- 分栏与窗口快捷键 ----------------
      // 切换双栏：Ctrl+\
      if (event.ctrlKey && !event.shiftKey && !event.altKey && (event.key === "\\" || code === "Backslash")) {
        event.preventDefault();
        useAppStore.getState().toggleSplit();
        return;
      }

      // 标签移到另一栏：Ctrl+Alt+← / →
      if (event.ctrlKey && event.altKey && !event.shiftKey && (event.key === "ArrowLeft" || event.key === "ArrowRight")) {
        event.preventDefault();
        const currentActiveId = useAppStore.getState().activeId;
        if (currentActiveId) {
          useAppStore.getState().moveDoc(currentActiveId, event.key === "ArrowLeft" ? 0 : 1);
        }
        return;
      }

      // ---------------- 编辑类快捷键 ----------------
      const editable = editorReady();
      const mdEditable = markdownReady();

      // 标题：Ctrl+1..6 设置级别，Ctrl+0 取消标题
      if (mdEditable && !event.shiftKey && !event.altKey && /^[0-6]$/.test(key)) {
        event.preventDefault();
        setHeading(key === "0" ? 0 : Number(key));
        return;
      }
      // 标题升降级：Ctrl+Alt+↑ / ↓
      if (mdEditable && event.altKey && (event.key === "ArrowUp" || event.key === "ArrowDown")) {
        event.preventDefault();
        shiftHeading(event.key === "ArrowUp" ? -1 : 1);
        return;
      }
      // 上标 / 下标 / 下划线 / 删除线：Alt+Shift+…
      // （Shift 会改变 event.key，例如 = 变 +、- 变 _，所以用 code 判断）
      if (mdEditable && event.altKey && event.shiftKey) {
        if (code === "Equal") {
          event.preventDefault();
          toggleSuperscript();
          return;
        }
        if (code === "Minus") {
          event.preventDefault();
          toggleSubscript();
          return;
        }
        if (code === "Digit5") {
          event.preventDefault();
          toggleStrikethrough();
          return;
        }
        if (code === "KeyU") {
          event.preventDefault();
          toggleUnderline();
          return;
        }
      }
      // 列表：Ctrl+Shift+7/8/9；高亮：Ctrl+Shift+M
      // （同样受 Shift 影响 event.key，用 code 判断）
      if (mdEditable && !event.altKey && event.shiftKey) {
        if (code === "Digit7") {
          event.preventDefault();
          toggleList("ordered");
          return;
        }
        if (code === "Digit8") {
          event.preventDefault();
          toggleList("bullet");
          return;
        }
        if (code === "Digit9") {
          event.preventDefault();
          toggleList("task");
          return;
        }
        if (code === "KeyM") {
          event.preventDefault();
          toggleHighlight();
          return;
        }
      }

      switch (key) {
        case "tab": {
          event.preventDefault();
          const { docs, activeId, activateDoc } = useAppStore.getState();
          if (docs.length < 2) break;
          const idx = docs.findIndex((d) => d.id === activeId);
          const next = docs[(idx + (event.shiftKey ? docs.length - 1 : 1)) % docs.length];
          if (next) activateDoc(next.id);
          break;
        }
        case "b": {
          if (!mdEditable || event.shiftKey || event.altKey) break;
          event.preventDefault();
          toggleBold();
          break;
        }
        case "i": {
          if (!mdEditable || event.altKey) break;
          event.preventDefault();
          if (event.shiftKey) {
            void import("../components/Toolbar/Toolbar").then((m) => m.promptInsert("image"));
          } else {
            toggleItalic();
          }
          break;
        }
        case "k": {
          if (event.shiftKey) {
            if (!editable) break;
            event.preventDefault();
            deleteLines();
          } else {
            if (!mdEditable || event.altKey) break;
            event.preventDefault();
            void import("../components/Toolbar/Toolbar").then((m) => m.promptInsert("link"));
          }
          break;
        }
        case "`": {
          if (!mdEditable || event.shiftKey) break;
          event.preventDefault();
          toggleInlineCode();
          break;
        }
        case "=":
        case "+": {
          event.preventDefault();
          const activeDoc = getActiveDoc();
          if (activeDoc) {
            const base = activeDoc.fontSize ?? useSettingsStore.getState().fontSize;
            useAppStore.getState().patchDoc(activeDoc.id, { fontSize: Math.min(32, base + 1) });
          }
          break;
        }
        case "-": {
          event.preventDefault();
          const activeDoc = getActiveDoc();
          if (activeDoc) {
            const base = activeDoc.fontSize ?? useSettingsStore.getState().fontSize;
            useAppStore.getState().patchDoc(activeDoc.id, { fontSize: Math.max(10, base - 1) });
          }
          break;
        }
        case "q": {
          if (!event.shiftKey || !mdEditable) break;
          event.preventDefault();
          toggleQuote();
          break;
        }
        case "l": {
          if (!event.shiftKey || !mdEditable) break;
          event.preventDefault();
          insertCallout("note");
          break;
        }
        case "c": {
          if (!event.shiftKey || !mdEditable) break;
          event.preventDefault();
          insertCodeBlock();
          break;
        }
        case "t": {
          if (!event.shiftKey || !mdEditable) break;
          event.preventDefault();
          insertTable(3, 3);
          break;
        }
        case "h": {
          if (!event.shiftKey || !mdEditable) break;
          event.preventDefault();
          insertHorizontalRule();
          break;
        }
        case "o": {
          if (!event.shiftKey || !mdEditable) break;
          event.preventDefault();
          insertToc();
          break;
        }
        case "d": {
          if (!event.shiftKey || !editable) break;
          event.preventDefault();
          duplicateLines();
          break;
        }
        default:
          break;
      }
    };

    // 使用捕获阶段，保证应用快捷键优先于 CodeMirror 的内部按键绑定
    window.addEventListener("keydown", handler, true);
    return () => window.removeEventListener("keydown", handler, true);
  }, []);
}

/** 监听当前文档变化，更新窗口标题 */
export function useWindowTitle(): void {
  const doc = useAppStore((s) => s.docs.find((d) => d.id === s.activeId) ?? null);
  const docs = useAppStore((s) => s.docs);
  const name = doc ? getDocTitle(doc, docs) : "未命名";
  const title = doc ? `${doc.isDirty ? "● " : ""}${name} - ${APP_NAME}` : APP_NAME;

  useEffect(() => {
    document.title = title;
    void (async () => {
      try {
        await getCurrentWindow().setTitle(title);
      } catch {
        /* 非 Tauri 环境忽略 */
      }
    })();
  }, [title]);
}

/** 把脏标记同步给 Rust，用于关闭窗口时判断是否需要询问 */
export function useDirtyFlagSync(): void {
  const dirty = useAppStore((s) => s.docs.some((d) => d.isDirty));
  useEffect(() => {
    void (async () => {
      try {
        await invoke("set_dirty", { dirty });
      } catch {
        /* ignore */
      }
    })();
  }, [dirty]);
}

/** 提供给外部读取当前活动文档 */
export function currentDoc() {
  return getActiveDoc();
}
