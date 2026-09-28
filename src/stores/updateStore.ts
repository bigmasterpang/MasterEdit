import { create } from "zustand";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { tempDir, join } from "@tauri-apps/api/path";
import { checkForUpdate, type UpdateInfo } from "../utils/updateCheck";

interface DownloadResult {
  path: string;
  size: number;
  sha256: string;
}

interface ProgressPayload {
  received: number;
  total: number;
}

interface UpdateState {
  info: UpdateInfo | null;
  checking: boolean;
  error: string | null;
  /** 是否已提示过（避免自动检查反复弹窗） */
  notified: boolean;
  dialogVisible: boolean;

  /** 下载状态 */
  downloading: boolean;
  progress: number;
  received: number;
  total: number;
  downloadedPath: string | null;
  installed: boolean;
  /** 正在替换并重启 */
  restarting: boolean;

  check: (options?: { silent?: boolean }) => Promise<void>;
  showDialog: () => void;
  hideDialog: () => void;
  downloadAndInstall: () => Promise<void>;
  /** 便携版：替换并重启 */
  applyPortableUpdate: () => Promise<void>;
  /** 安装版：静默升级并重启 */
  applyInstallerUpdate: () => Promise<void>;
  runInstaller: () => Promise<void>;
  resetDownload: () => void;
}

export const useUpdateStore = create<UpdateState>((set, get) => ({
  info: null,
  checking: false,
  error: null,
  notified: false,
  dialogVisible: false,

  downloading: false,
  progress: 0,
  received: 0,
  total: 0,
  downloadedPath: null,
  installed: false,
  restarting: false,

  check: async (options = {}) => {
    if (get().checking) return;
    set({ checking: true, error: null });
    try {
      const info = await checkForUpdate();
      set({ info, checking: false });
      if (!options.silent) {
        // 用户主动点击：始终弹出对话框显示结果（修复第二次点击不弹窗的 BUG）
        set({ dialogVisible: true });
      } else if (info.hasUpdate && !get().notified) {
        // 后台静默自动检查：只在有更新且未提示过时弹窗
        set({ notified: true, dialogVisible: true });
      }
    } catch (error) {
      set({ checking: false, error: String(error) });
      if (!options.silent) set({ dialogVisible: true });
    }
  },

  showDialog: () => set({ dialogVisible: true }),
  hideDialog: () => set({ dialogVisible: false }),

  /** 下载安装包（边下边校验 SHA-256），完成后自动启动安装程序 */
  downloadAndInstall: async () => {
    const info = get().info;
    if (!info?.downloadUrl) {
      set({ error: "当前更新源不支持应用内下载，请打开发布页面手动下载。" });
      return;
    }
    set({
      downloading: true,
      progress: 0,
      received: 0,
      total: info.size ?? 0,
      error: null,
      downloadedPath: null,
      installed: false,
    });

    const unlisten = await listen<ProgressPayload>("update-progress", (event) => {
      const { received, total } = event.payload;
      set({
        received,
        total: total || received,
        progress: total > 0 ? Math.min(1, received / total) : 0,
      });
    });

    try {
      const dir = await tempDir();
      const filename = info.filename || `MasterEdit_${info.latest}_x64-setup.exe`;
      const dest = await join(dir, filename);

      const result = await invoke<DownloadResult>("download_update", {
        url: info.downloadUrl,
        dest,
        expectSha256: info.sha256 ?? null,
        expectSize: info.size ?? null,
      });

      set({
        downloading: false,
        progress: 1,
        downloadedPath: result.path,
      });

      // 根据安装类型分流：安装版走静默安装脚本；若无标记但文件名匹配 setup.exe 兼容走向导；绿色版走自替换
      if (info.installKind === "installed" || info.installKind === "installer") {
        await get().applyInstallerUpdate();
      } else if (/setup\.exe$/i.test(info.filename ?? "")) {
        await get().runInstaller();
      } else {
        await get().applyPortableUpdate();
      }
    } catch (error) {
      set({ downloading: false, error: String(error) });
    } finally {
      unlisten();
    }
  },

  /** 便携版自更新：替换 exe 后由新进程接管，当前进程退出 */
  applyPortableUpdate: async () => {
    const path = get().downloadedPath;
    if (!path) return;
    try {
      set({ restarting: true });
      await invoke("apply_update", { path });
      set({ installed: true });
      // 给界面一点时间显示"正在重启"，随后强制退出旧进程
      await new Promise((resolve) => setTimeout(resolve, 700));
      try {
        await invoke("quit_app");
      } catch {
        await invoke("confirm_close");
      }
    } catch (error) {
      set({ restarting: false, error: String(error) });
    }
  },

  /** 安装版静默升级：调用临时 PowerShell 脚本在当前进程退出后静默运行 setup.exe /S 并重启 */
  applyInstallerUpdate: async () => {
    const path = get().downloadedPath;
    if (!path) return;
    try {
      set({ restarting: true });
      await invoke("apply_installer_update", { path });
      set({ installed: true });
      await new Promise((resolve) => setTimeout(resolve, 700));
      try {
        await invoke("quit_app");
      } catch {
        await invoke("confirm_close");
      }
    } catch (error) {
      set({ restarting: false, error: String(error) });
    }
  },

  runInstaller: async () => {
    const path = get().downloadedPath;
    if (!path) return;
    try {
      await invoke("run_installer", { path });
      set({ installed: true });
    } catch (error) {
      set({ error: String(error) });
    }
  },

  resetDownload: () =>
    set({
      downloadedPath: null,
      progress: 0,
      received: 0,
      installed: false,
      restarting: false,
    }),
}));
