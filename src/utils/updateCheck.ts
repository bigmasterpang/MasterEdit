import { getVersion } from "@tauri-apps/api/app";
import { invoke } from "@tauri-apps/api/core";
import { openUrl } from "@tauri-apps/plugin-opener";

export const GITHUB_REPO = "bigmasterpang/MasterEdit";
export const RELEASES_URL = `https://github.com/${GITHUB_REPO}/releases`;
export const PORTAL_URL = "https://master.dapang.wang";

/** 软件中心返回的最新版本信息 */
export interface PortalRelease {
  baseUrl: string;
  version: string;
  filename: string;
  size: number;
  humanSize: string;
  sha256: string;
  uploadedAt: string;
  releaseNotes: string;
  downloadUrl: string;
  hasUpdate: boolean;
  currentVersion: string;
  installKind: "portable" | "installed" | "installer";
}

export interface UpdateInfo {
  current: string;
  latest: string | null;
  hasUpdate: boolean;
  publishedAt: string | null;
  notes: string;
  /** 详情页地址（软件中心或 GitHub Release） */
  url: string;
  noRelease: boolean;
  checkedAt: number;
  /** 更新来源：软件中心（可内下载安装）或 GitHub（仅跳转） */
  source: "portal" | "github";
  size?: number;
  humanSize?: string;
  sha256?: string;
  downloadUrl?: string;
  filename?: string;
  /** 安装类型：绿色便携版 / NSIS 安装版 */
  installKind?: "portable" | "installed" | "installer";
}

interface GithubRelease {
  tag_name?: string;
  name?: string;
  body?: string;
  html_url?: string;
  published_at?: string;
}

interface GithubTag {
  name?: string;
}

/** 版本号比较：返回 1 表示 a 比 b 新 */
export function compareVersion(a: string, b: string): number {
  const normalize = (value: string) =>
    value
      .trim()
      .replace(/^v/i, "")
      .split(/[.+-]/)
      .map((part) => Number.parseInt(part, 10))
      .map((n) => (Number.isFinite(n) ? n : 0));
  const left = normalize(a);
  const right = normalize(b);
  const length = Math.max(left.length, right.length);
  for (let i = 0; i < length; i++) {
    const l = left[i] ?? 0;
    const r = right[i] ?? 0;
    if (l > r) return 1;
    if (l < r) return -1;
  }
  return 0;
}

/** GitHub 兜底查询（接口支持 CORS，可直接 fetch） */
async function checkGithub(current: string): Promise<UpdateInfo | null> {
  const fetchJson = async <T>(url: string): Promise<T | null> => {
    try {
      const response = await fetch(url, {
        headers: { Accept: "application/vnd.github+json" },
        cache: "no-store",
      });
      if (!response.ok) return null;
      return (await response.json()) as T;
    } catch {
      return null;
    }
  };

  const release = await fetchJson<GithubRelease>(
    `https://api.github.com/repos/${GITHUB_REPO}/releases/latest`,
  );
  if (release?.tag_name) {
    return {
      current,
      latest: release.tag_name,
      hasUpdate: compareVersion(release.tag_name, current) > 0,
      publishedAt: release.published_at ?? null,
      notes: release.body ?? "",
      url: release.html_url ?? RELEASES_URL,
      noRelease: false,
      checkedAt: Date.now(),
      source: "github",
    };
  }

  const tags = await fetchJson<GithubTag[]>(
    `https://api.github.com/repos/${GITHUB_REPO}/tags?per_page=1`,
  );
  const tag = Array.isArray(tags) && tags.length > 0 ? tags[0].name ?? null : null;
  if (tag) {
    return {
      current,
      latest: tag,
      hasUpdate: compareVersion(tag, current) > 0,
      publishedAt: null,
      notes: "",
      url: RELEASES_URL,
      noRelease: false,
      checkedAt: Date.now(),
      source: "github",
    };
  }
  return null;
}

/** 前端直接请求软件中心双节点（当 Rust WinHTTP 遇特殊环境异常时的二级保障） */
async function checkPortalDirect(current: string): Promise<UpdateInfo | null> {
  const nodes = ["https://master.dapang.wang", "http://106.14.225.57"];
  for (const base of nodes) {
    for (const variant of ["installer", "portable"]) {
      try {
        const resp = await fetch(`${base}/api/apps/masteredit/windows/latest?variant=${variant}`, {
          cache: "no-store",
        });
        if (resp.ok) {
          const json = await resp.json() as Record<string, unknown>;
          const version = typeof json.version === "string" ? json.version : "";
          if (version) {
            const dl = typeof json.download_url === "string" ? json.download_url : "";
            return {
              current,
              latest: version,
              hasUpdate: compareVersion(version, current) > 0,
              publishedAt: typeof json.uploaded_at === "string" ? json.uploaded_at : null,
              notes: typeof json.release_notes === "string" ? json.release_notes : "",
              url: `${base}${dl}`,
              noRelease: false,
              checkedAt: Date.now(),
              source: "portal",
              size: typeof json.size === "number" ? json.size : undefined,
              humanSize: typeof json.human_size === "string" ? json.human_size : undefined,
              sha256: typeof json.sha256 === "string" ? json.sha256 : undefined,
              downloadUrl: `${base}${dl}`,
              filename: typeof json.filename === "string" ? json.filename : undefined,
              installKind: (variant === "installer" ? "installed" : "portable") as "portable" | "installed",
            };
          }
        }
      } catch {
        /* try next */
      }
    }
  }
  return null;
}

/**
 * 检查更新：优先软件中心（master.dapang.wang → 国内节点回退），
 * 软件中心彻底不可达时才回退到 GitHub Release。
 */
export async function checkForUpdate(): Promise<UpdateInfo> {
  const current = await getVersion().catch(() => "0.0.0");
  let portalError = "";

  try {
    const release = await invoke<PortalRelease>("check_update");
    return {
      current: release.currentVersion || current,
      latest: release.version,
      hasUpdate: release.hasUpdate,
      publishedAt: release.uploadedAt || null,
      notes: release.releaseNotes || "",
      url: release.downloadUrl || PORTAL_URL,
      noRelease: false,
      checkedAt: Date.now(),
      source: "portal",
      size: release.size,
      humanSize: release.humanSize,
      sha256: release.sha256,
      downloadUrl: release.downloadUrl,
      filename: release.filename,
      installKind: release.installKind ?? "portable",
    };
  } catch (error) {
    portalError = String(error);
  }

  // 二级保障：若 Rust 内部异常，前端直接通过 fetch 访问软件中心
  const directPortal = await checkPortalDirect(current);
  if (directPortal) {
    return directPortal;
  }

  // 兜底：仅在软件中心彻底无法连接时才查 GitHub
  const github = await checkGithub(current);
  if (github) return { ...github, installKind: "portable" };

  const offline: UpdateInfo = {
    current,
    latest: null,
    hasUpdate: false,
    publishedAt: null,
    notes: "",
    url: PORTAL_URL,
    noRelease: true,
    checkedAt: Date.now(),
    source: "portal",
    installKind: "portable",
  };
  if (portalError.includes("暂无")) return offline;
  throw new Error(portalError || "无法连接更新服务器");
}

export async function openReleasesPage(url = RELEASES_URL): Promise<void> {
  try {
    await openUrl(url);
  } catch {
    /* ignore */
  }
}

export function formatSize(bytes?: number, fallback?: string): string {
  if (fallback) return fallback;
  if (!bytes) return "-";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}
