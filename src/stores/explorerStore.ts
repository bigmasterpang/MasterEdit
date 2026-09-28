import { create } from "zustand";
import { invoke } from "@tauri-apps/api/core";
import { normalizeSlashes, samePath } from "../utils/filePath";
import { debounce } from "../utils/timing";
import { readMigratedItem } from "../utils/storage";
import type { DocState } from "../types";
import { getActiveDoc } from "./appStore";

export interface DirEntry {
  name: string;
  path: string;
  isDir: boolean;
  size: number;
  modifiedAt: number;
  ext: string;
}

type ExplorerFilter = "openable" | "all";

interface ExplorerState {
  visible: boolean;
  rootPath: string | null;
  children: Record<string, DirEntry[]>;
  loading: Record<string, boolean>;
  expanded: Record<string, boolean>;
  filter: ExplorerFilter;
  showHidden: boolean;
  ratio: number;
  sidebarWidth: number;
  tabRoots: Record<string, string>;

  setVisible: (v: boolean) => void;
  toggle: () => void;
  setRoot: (path: string) => Promise<void>;
  setDocDirectory: (docFilePath: string | null) => void;
  setSidebarWidth: (w: number) => void;
  bindDocRoot: (docId: string, folderPath: string) => void;
  syncToDoc: (doc: DocState | null) => Promise<void>;
  toggleDir: (path: string) => Promise<void>;
  expandTo: (targetFilePath: string) => Promise<void>;
  refresh: (path?: string) => Promise<void>;
  setFilter: (f: ExplorerFilter) => void;
  setShowHidden: (v: boolean) => void;
  setRatio: (r: number) => void;
}

/** localStorage 持久化 key（旧键 mastermd.explorer 由 readMigratedItem 兼容读取） */
const STORAGE_KEY = "masteredit.explorer";
const LEGACY_STORAGE_KEY = "mastermd.explorer";
/** expanded 最多持久化的条数，超出截断 */
const EXPANDED_LIMIT = 200;
/** 上下两栏比例范围与默认值 */
const MIN_RATIO = 0.2;
const MAX_RATIO = 0.8;
const DEFAULT_RATIO = 0.55;
/** 侧边栏宽度限制与默认值 */
const MIN_SIDEBAR_WIDTH = 180;
const MAX_SIDEBAR_WIDTH = 600;
const DEFAULT_SIDEBAR_WIDTH = 260;
/** 父目录链最大深度，防止异常路径导致死循环 */
const MAX_DEPTH = 64;

interface PersistedExplorer {
  rootPath: string | null;
  ratio: number;
  sidebarWidth: number;
  showHidden: boolean;
  filter: ExplorerFilter;
  expanded: Record<string, boolean>;
  tabRoots: Record<string, string>;
}

/** 当前根目录是否由文档目录自动推导：自动根跟随文档切换，用户手选的根不覆盖 */
let rootFromDoc = false;

/** 限制侧边栏宽度在合理区间 */
function clampSidebarWidth(w: number): number {
  if (!Number.isFinite(w)) return DEFAULT_SIDEBAR_WIDTH;
  return Math.min(MAX_SIDEBAR_WIDTH, Math.max(MIN_SIDEBAR_WIDTH, Math.round(w)));
}

/** 收敛比例到 0.2 ~ 0.8 */
function clampRatio(r: number): number {
  if (!Number.isFinite(r)) return DEFAULT_RATIO;
  return Math.min(MAX_RATIO, Math.max(MIN_RATIO, r));
}

/** 去掉结尾分隔符（保留盘符根 `C:\` 与 UNC 前缀） */
function stripTail(p: string): string {
  if (/^[a-zA-Z]:\\$/.test(p)) return p;
  return p.replace(/[\\/]+$/, "");
}

/** 取父目录（Windows 风格，保留 UNC 双反斜杠）；无父级返回空串 */
export function parentPath(p: string): string {
  const n = stripTail(p);
  const idx = Math.max(n.lastIndexOf("\\"), n.lastIndexOf("/"));
  if (idx < 0) return "";
  if (idx === 0) return "\\";
  const parent = n.slice(0, idx);
  // 盘符根：C: -> C:\
  if (/^[a-zA-Z]:$/.test(parent)) return `${parent}\\`;
  return parent;
}

/** 判断 child 是否位于 ancestor 之下（含相等，忽略大小写与分隔符差异） */
function isUnder(child: string, ancestor: string): boolean {
  const c = normalizeSlashes(stripTail(child)).toLowerCase();
  const a = normalizeSlashes(stripTail(ancestor)).toLowerCase();
  if (!a) return false;
  const prefix = a.endsWith("\\") ? a : `${a}\\`;
  return c === a || c.startsWith(prefix);
}

/** 读取持久化数据；任何异常都静默回退默认值 */
function loadPersisted(): Partial<PersistedExplorer> {
  try {
    const raw = readMigratedItem(STORAGE_KEY, LEGACY_STORAGE_KEY);
    if (!raw) return {};
    const data = JSON.parse(raw) as Partial<PersistedExplorer>;
    const out: Partial<PersistedExplorer> = {};
    if (typeof data.rootPath === "string" && data.rootPath) out.rootPath = data.rootPath;
    if (typeof data.ratio === "number") out.ratio = clampRatio(data.ratio);
    if (typeof data.sidebarWidth === "number") out.sidebarWidth = clampSidebarWidth(data.sidebarWidth);
    if (typeof data.showHidden === "boolean") out.showHidden = data.showHidden;
    if (data.filter === "openable" || data.filter === "all") out.filter = data.filter;
    if (data.expanded && typeof data.expanded === "object") {
      out.expanded = Object.fromEntries(
        Object.entries(data.expanded)
          .filter(([, value]) => value === true)
          .slice(0, EXPANDED_LIMIT),
      );
    }
    if (data.tabRoots && typeof data.tabRoots === "object") {
      out.tabRoots = { ...data.tabRoots };
    }
    return out;
  } catch (error) {
    console.error("读取资源管理器状态失败，使用默认值", error);
    return {};
  }
}

const persisted = loadPersisted();

/** 提取需要持久化的字段（expanded 截断到 200 条） */
function pickPersisted(state: ExplorerState): PersistedExplorer {
  return {
    rootPath: state.rootPath,
    ratio: state.ratio,
    sidebarWidth: state.sidebarWidth,
    showHidden: state.showHidden,
    filter: state.filter,
    expanded: Object.fromEntries(
      Object.entries(state.expanded)
        .filter(([, value]) => value)
        .slice(0, EXPANDED_LIMIT),
    ),
    tabRoots: state.tabRoots,
  };
}

/** 防抖写入，避免拖拽 / 连续展开时频繁写 localStorage */
const schedulePersist = debounce(() => {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify(pickPersisted(useExplorerStore.getState())),
    );
  } catch (error) {
    console.error("保存资源管理器状态失败", error);
  }
}, 300);

/** 加载目录内容；失败时写入空数组并记录日志，不抛出（避免打断 UI） */
async function loadDir(path: string, force = false): Promise<void> {
  const state = useExplorerStore.getState();
  if (!force && state.children[path] !== undefined) return;
  if (state.loading[path]) return;
  useExplorerStore.setState({ loading: { ...state.loading, [path]: true } });
  try {
    const list = await invoke<DirEntry[]>("list_directory", {
      path,
      showHidden: state.showHidden,
    });
    const current = useExplorerStore.getState();
    useExplorerStore.setState({
      children: { ...current.children, [path]: list },
      loading: { ...current.loading, [path]: false },
    });
  } catch (error) {
    console.error(`加载目录失败: ${path}`, error);
    const current = useExplorerStore.getState();
    useExplorerStore.setState({
      children: { ...current.children, [path]: [] },
      loading: { ...current.loading, [path]: false },
    });
  }
}

/** 隐藏项开关变化后，重新拉取已加载过的根目录与展开目录 */
async function reloadLoaded(): Promise<void> {
  const { rootPath, expanded, children } = useExplorerStore.getState();
  const targets = new Set<string>();
  if (rootPath && children[rootPath] !== undefined) targets.add(rootPath);
  for (const path of Object.keys(children)) {
    if (expanded[path]) targets.add(path);
  }
  for (const path of targets) {
    await loadDir(path, true);
  }
}

/** 面板首次可见时恢复：加载根目录，并补载已记住的展开目录 */
async function hydrateVisible(): Promise<void> {
  const { rootPath, children } = useExplorerStore.getState();
  if (!rootPath) return;
  if (children[rootPath] === undefined) await loadDir(rootPath);
  // 逐层补载：父目录已加载且自身仍处于展开状态的目录
  for (let round = 0; round < MAX_DEPTH; round += 1) {
    const state = useExplorerStore.getState();
    const pending = Object.keys(state.expanded).filter(
      (path) =>
        path !== rootPath &&
        state.expanded[path] &&
        isUnder(path, rootPath) &&
        state.children[path] === undefined &&
        state.children[parentPath(path)] !== undefined,
    );
    if (pending.length === 0) break;
    for (const path of pending) await loadDir(path);
  }
}

export const useExplorerStore = create<ExplorerState>((set, get) => ({
  visible: false,
  rootPath: persisted.rootPath ?? null,
  children: {},
  loading: {},
  expanded: persisted.expanded ?? {},
  filter: persisted.filter ?? "openable",
  showHidden: persisted.showHidden ?? false,
  sidebarWidth: persisted.sidebarWidth ?? DEFAULT_SIDEBAR_WIDTH,
  tabRoots: persisted.tabRoots ?? {},
  ratio: persisted.ratio ?? DEFAULT_RATIO,

  setVisible: (v) => {
    set({ visible: v });
    // 恢复的根目录在面板首次可见时按需加载，避免启动时无谓 IO
    if (v) void hydrateVisible();
  },

  toggle: () => get().setVisible(!get().visible),

  setSidebarWidth: (w) => {
    set({ sidebarWidth: clampSidebarWidth(w) });
    schedulePersist();
  },

  bindDocRoot: (docId, folderPath) => {
    const clean = stripTail(folderPath);
    if (!clean) return;
    set((s) => ({
      tabRoots: { ...s.tabRoots, [docId]: clean },
    }));
    schedulePersist();
  },

  syncToDoc: async (doc) => {
    if (!doc) return;
    const { tabRoots, rootPath } = get();
    const bound = tabRoots[doc.id];
    if (bound) {
      if (!rootPath || !samePath(rootPath, bound)) {
        await get().setRoot(bound);
      }
      if (doc.filePath) {
        await get().expandTo(doc.filePath);
      }
    } else if (doc.filePath) {
      const dir = parentPath(doc.filePath);
      if (dir) {
        get().bindDocRoot(doc.id, dir);
        if (!rootPath || !samePath(rootPath, dir)) {
          await get().setRoot(dir);
        }
        await get().expandTo(doc.filePath);
      }
    }
  },

  setRoot: async (path) => {
    const root = stripTail(path);
    if (!root) return;
    rootFromDoc = false;
    const activeDoc = getActiveDoc();
    const nextTabRoots = activeDoc
      ? { ...get().tabRoots, [activeDoc.id]: root }
      : get().tabRoots;
    set({
      rootPath: root,
      children: {},
      loading: {},
      expanded: { ...get().expanded, [root]: true },
      tabRoots: nextTabRoots,
    });
    schedulePersist();
    await loadDir(root);
  },

  setDocDirectory: (docFilePath) => {
    if (!docFilePath) return;
    const dir = parentPath(docFilePath);
    if (!dir) return;
    const { rootPath } = get();
    // 用户手选的根目录不因文档切换而改变；自动推导的根则跟随当前文档
    if (rootPath && !rootFromDoc) return;
    if (rootPath && samePath(rootPath, dir)) return;
    void get().setRoot(dir);
    // setRoot 会把手选标记置否，这里恢复为「跟随文档」
    rootFromDoc = true;
  },

  toggleDir: async (path) => {
    const target = stripTail(path);
    if (!target) return;
    const { expanded, children } = get();
    const next = !expanded[target];
    set({ expanded: { ...expanded, [target]: next } });
    schedulePersist();
    // 首次展开才加载（懒加载）
    if (next && children[target] === undefined) {
      await loadDir(target);
    }
  },

  expandTo: async (targetFilePath) => {
    const target = stripTail(targetFilePath);
    if (!target) return;
    const dir = parentPath(target);
    if (!dir) return;

    const root = get().rootPath;
    // 目标不在当前根下：把根切到文件所在目录，文件即可直接可见
    if (!root || !isUnder(target, root)) {
      await get().setRoot(dir);
      rootFromDoc = true;
      return;
    }

    // 从文件所在目录向上收集到根（不含根）的目录链
    const chain: string[] = [];
    let cursor = dir;
    let depth = 0;
    while (cursor && !samePath(cursor, root) && depth < MAX_DEPTH) {
      chain.unshift(cursor);
      const parent = await invoke<string | null>("parent_dir_of", { path: cursor });
      if (!parent || samePath(parent, cursor)) break;
      cursor = stripTail(parent);
      depth += 1;
    }

    // 先展开根，再自顶向下逐级展开并加载，确保目标文件可见
    const rootState = get();
    if (!rootState.expanded[root]) {
      set({ expanded: { ...rootState.expanded, [root]: true } });
      schedulePersist();
    }
    await loadDir(root);
    for (const path of chain) {
      const state = get();
      if (!state.expanded[path]) {
        set({ expanded: { ...state.expanded, [path]: true } });
        schedulePersist();
      }
      await loadDir(path);
    }
  },

  refresh: async (path) => {
    const target = path ? stripTail(path) : get().rootPath;
    if (!target) return;
    await loadDir(target, true);
  },

  setFilter: (f) => {
    set({ filter: f });
    schedulePersist();
  },

  setShowHidden: (v) => {
    if (get().showHidden === v) return;
    set({ showHidden: v });
    schedulePersist();
    // 已加载目录按新规则重新拉取，避免缓存里残留隐藏项
    void reloadLoaded();
  },

  setRatio: (r) => {
    set({ ratio: clampRatio(r) });
    schedulePersist();
  },
}));
