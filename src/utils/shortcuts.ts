import type { ShortcutId, ShortcutMap } from "../types";

export const SHORTCUT_IDS: ShortcutId[] = [
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

export const SHORTCUT_LABELS: Record<ShortcutId, string> = {
  open: "打开文件",
  save: "保存",
  saveAs: "另存为",
  newDoc: "新建文档",
  viewMode: "切换视图模式",
  search: "查找",
  replace: "查找并替换",
  closeTab: "关闭当前标签",
  settings: "打开设置",
};

export const DEFAULT_SHORTCUTS: ShortcutMap = {
  open: "Ctrl+O",
  save: "Ctrl+S",
  saveAs: "Ctrl+Shift+S",
  newDoc: "Ctrl+N",
  viewMode: "Ctrl+E",
  search: "Ctrl+F",
  replace: "Ctrl+H",
  closeTab: "Ctrl+W",
  settings: "Ctrl+,",
};

export interface ParsedShortcut {
  ctrl: boolean;
  shift: boolean;
  alt: boolean;
  key: string;
}

/** 解析 "Ctrl+Shift+S" 形式的快捷键 */
export function parseShortcut(text: string): ParsedShortcut | null {
  if (!text) return null;
  const parts = text
    .split("+")
    .map((p) => p.trim())
    .filter(Boolean);
  if (parts.length === 0) return null;
  const parsed: ParsedShortcut = { ctrl: false, shift: false, alt: false, key: "" };
  for (const part of parts) {
    const lower = part.toLowerCase();
    if (lower === "ctrl" || lower === "control") parsed.ctrl = true;
    else if (lower === "shift") parsed.shift = true;
    else if (lower === "alt") parsed.alt = true;
    else parsed.key = lower;
  }
  if (!parsed.key) return null;
  return parsed;
}

/** 键盘事件是否匹配某个快捷键 */
export function matchesShortcut(event: KeyboardEvent, shortcut: string): boolean {
  const parsed = parseShortcut(shortcut);
  if (!parsed) return false;
  const ctrl = event.ctrlKey || event.metaKey;
  if (ctrl !== parsed.ctrl) return false;
  if (event.shiftKey !== parsed.shift) return false;
  if (event.altKey !== parsed.alt) return false;
  return event.key.toLowerCase() === parsed.key;
}

/** 键盘事件转快捷键字符串（用于自定义录入） */
export function eventToShortcut(event: KeyboardEvent): string | null {
  const key = event.key;
  if (["Control", "Shift", "Alt", "Meta", "CapsLock", "Escape", "Tab"].includes(key)) {
    return null;
  }
  if (!(event.ctrlKey || event.metaKey)) return null;
  const parts: string[] = ["Ctrl"];
  if (event.shiftKey) parts.push("Shift");
  if (event.altKey) parts.push("Alt");
  const normalized = key.length === 1 ? key.toUpperCase() : key;
  parts.push(normalized);
  return parts.join("+");
}

/** 检查快捷键冲突，返回冲突的其它功能 id */
export function findConflict(
  map: ShortcutMap,
  id: ShortcutId,
  value: string,
): ShortcutId | null {
  const target = value.toLowerCase();
  for (const other of SHORTCUT_IDS) {
    if (other === id) continue;
    if ((map[other] ?? "").toLowerCase() === target) return other;
  }
  return null;
}

export interface ShortcutReferenceGroup {
  title: string;
  items: Array<{ keys: string; label: string; configurable?: boolean }>;
}

/** 应用内「全部快捷键」参考表（设置面板展示） */
export const SHORTCUT_REFERENCE: ShortcutReferenceGroup[] = [
  {
    title: "文件",
    items: [
      { keys: "Ctrl+O", label: "打开文件", configurable: true },
      { keys: "Ctrl+S", label: "保存", configurable: true },
      { keys: "Ctrl+Shift+S", label: "另存为", configurable: true },
      { keys: "Ctrl+N", label: "新建文档", configurable: true },
      { keys: "Ctrl+W", label: "关闭当前标签", configurable: true },
    ],
  },
  {
    title: "表格编辑（CSV / xlsx）",
    items: [
      { keys: "双击 / F2 / 直接输入", label: "编辑当前单元格（中文输入法请先双击或 F2）" },
      { keys: "Enter / Tab", label: "提交并下移 / 右移（Shift+Tab 左移）" },
      { keys: "Esc", label: "取消本次编辑" },
      { keys: "Ctrl+Z / Ctrl+Y", label: "撤销 / 重做单元格编辑" },
      { keys: "Ctrl+C", label: "复制当前单元格文本" },
      { keys: "Ctrl+E", label: "CSV：表格 / 源码 / 分屏 切换" },
      { keys: "Ctrl+S", label: "写回文件（xlsx 保存前会备份 .bak）" },
    ],
  },
  {
    title: "标题与文本格式",
    items: [
      { keys: "Ctrl+1 … Ctrl+6", label: "设置为 H1–H6 标题" },
      { keys: "Ctrl+0", label: "取消标题（变回普通段落）" },
      { keys: "Ctrl+Alt+↑ / Ctrl+Alt+↓", label: "标题提升 / 降低一级" },
      { keys: "Ctrl+B / Ctrl+I", label: "粗体 / 斜体" },
      { keys: "Ctrl+K", label: "插入链接（弹窗填写）" },
      { keys: "Ctrl+`", label: "行内代码" },
      { keys: "Alt+Shift+5", label: "删除线" },
      { keys: "Alt+Shift+= / Alt+Shift+-", label: "上标 / 下标" },
      { keys: "Ctrl+Shift+M", label: "高亮 ==文字==" },
      { keys: "Alt+Shift+U", label: "下划线 ++文字++" },
    ],
  },
  {
    title: "块级结构",
    items: [
      { keys: "Ctrl+Shift+7 / 8 / 9", label: "有序 / 无序 / 任务列表" },
      { keys: "Ctrl+Shift+Q", label: "引用块" },
      { keys: "Ctrl+Shift+L", label: "提示块（NOTE）" },
      { keys: "Ctrl+Shift+C", label: "代码块" },
      { keys: "Ctrl+Shift+T", label: "表格 3×3" },
      { keys: "Ctrl+Shift+H", label: "水平分割线" },
      { keys: "Ctrl+Shift+I", label: "插入图片（弹窗填写）" },
      { keys: "Ctrl+Shift+O", label: "插入目录 (TOC)" },
    ],
  },
  {
    title: "行操作",
    items: [
      { keys: "Alt+↑ / Alt+↓", label: "上移 / 下移当前行" },
      { keys: "Shift+Alt+↑ / Shift+Alt+↓", label: "复制当前行" },
      { keys: "Ctrl+Shift+D", label: "复制当前行（备用）" },
      { keys: "Ctrl+Shift+K", label: "删除当前行" },
    ],
  },
  {
    title: "代码查看与导航",
    items: [
      { keys: "F12 / Ctrl+左键", label: "转到原函数 / 符号定义" },
      { keys: "Alt+F12", label: "速览原函数实现 (Peek Definition)" },
      { keys: "Ctrl+Alt+左键", label: "在右侧分栏打开原函数定义" },
      { keys: "Shift+F12", label: "查找所有引用 (Find All References)" },
      { keys: "Alt+← / Alt+→", label: "代码跳转位置后退 / 前进（支持鼠标侧键）" },
    ],
  },
  {
    title: "视图与查找",
    items: [
      { keys: "Ctrl+E", label: "切换视图模式（预览/源码/分屏）", configurable: true },
      { keys: "Ctrl+\\", label: "开启 / 关闭左右双栏并排" },
      { keys: "Alt+1 / Alt+2", label: "聚焦左栏 / 右栏" },
      { keys: "Ctrl+F", label: "查找", configurable: true },
      { keys: "Ctrl+H", label: "查找并替换", configurable: true },
      { keys: "F3 / Shift+F3", label: "查找下一个 / 上一个" },
      { keys: "Enter / Shift+Enter", label: "（查找框内）下一个 / 上一个" },
      { keys: "Ctrl+Tab / Ctrl+Shift+Tab", label: "切换到下一个 / 上一个标签页" },
      { keys: "Ctrl+= / Ctrl+-", label: "当前窗口文档字号增大 / 减小" },
      { keys: "Ctrl+滚轮", label: "缩放当前窗口文档字号（双栏互不影响）" },
      { keys: "Ctrl+,", label: "打开设置", configurable: true },
      { keys: "F1", label: "快捷键面板" },
      { keys: "右键", label: "编辑器 / 预览区上下文菜单" },
      { keys: "Esc", label: "关闭搜索栏 / 弹窗" },
    ],
  },
];
