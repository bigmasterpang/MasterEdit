import { useEffect, useMemo, useState } from "react";
import { Modal, Button } from "../common/Modal";
import { useDialogStore } from "../../stores/dialogStore";
import { useUpdateStore } from "../../stores/updateStore";
import { useDungeonStore } from "../../stores/dungeonStore";
import { useAppStore } from "../../stores/appStore";
import { openReleasesPage, PORTAL_URL } from "../../utils/updateCheck";
import {
  APP_NAME,
  AUTHOR_NAME,
  AUTHOR_NAME_CN,
  PROJECT_URL,
} from "../../utils/constants";
import { getVersion } from "@tauri-apps/api/app";
import { readMigratedItem, writeItem } from "../../utils/storage";

const EGG_STORAGE_KEY = "masteredit.easter-egg";
/** 改名前的键名，读取时兼容迁移 */
const EGG_LEGACY_STORAGE_KEY = "mastermd.easter-egg";
const EGG_CLICKS = 7;

interface ConfettiPiece {
  id: number;
  left: number;
  delay: number;
  duration: number;
  color: string;
  rotate: number;
  width: number;
  height: number;
}

const CONFETTI_COLORS = [
  "#3B82F6",
  "#8B5CF6",
  "#F59E0B",
  "#10B981",
  "#EF4444",
  "#EC4899",
  "#06B6D4",
];

function makeConfetti(count = 42): ConfettiPiece[] {
  return Array.from({ length: count }, (_, index) => ({
    id: Date.now() + index,
    left: Math.random() * 100,
    delay: Math.random() * 0.5,
    duration: 1.6 + Math.random() * 1.4,
    color: CONFETTI_COLORS[index % CONFETTI_COLORS.length],
    rotate: Math.random() * 360,
    width: 6 + Math.random() * 6,
    height: 8 + Math.random() * 10,
  }));
}

/** 关于 MasterEdit：版本、作者、软件中心与隐藏彩蛋 */
export function AboutDialog() {
  const open = useDialogStore((s) => s.aboutVisible);
  const close = () => useDialogStore.getState().setAboutVisible(false);
  const info = useUpdateStore((s) => s.info);
  const checking = useUpdateStore((s) => s.checking);
  const [version, setVersion] = useState("");
  const [exePath, setExePath] = useState("");
  const [clicks, setClicks] = useState(0);
  const [eggUnlocked, setEggUnlocked] = useState(
    () => readMigratedItem(EGG_STORAGE_KEY, EGG_LEGACY_STORAGE_KEY) === "1",
  );
  const [confetti, setConfetti] = useState<ConfettiPiece[]>([]);
  const dungeonMeta = useDungeonStore((s) => s.meta);

  useEffect(() => {
    // 每次打开/关闭都重置临时状态（点击计数、提示语、彩纸）
    setClicks(0);
    setConfetti([]);
    if (!open) return;
    void getVersion()
      .then(setVersion)
      .catch(() => setVersion("0.0.0"));
    void import("@tauri-apps/api/core")
      .then(({ invoke }) => invoke<string | null>("current_exe_path"))
      .then((path) => setExePath(path ?? ""))
      .catch(() => setExePath(""));
  }, [open]);

  const hints = useMemo(() => {
    if (eggUnlocked) return "";
    if (clicks === 0) return "";
    if (clicks < 3) return "嗯？";
    if (clicks < 5) return "继续点…";
    if (clicks < EGG_CLICKS - 1) return "还差一点…";
    return "马上就好…";
  }, [clicks, eggUnlocked]);

  const handleLogoClick = () => {
    if (eggUnlocked) {
      setConfetti(makeConfetti());
      return;
    }
    const next = clicks + 1;
    setClicks(next);
    if (next >= EGG_CLICKS) {
      writeItem(EGG_STORAGE_KEY, "1");
      setEggUnlocked(true);
      setConfetti(makeConfetti(60));
    }
  };

  return (
    <Modal
      open={open}
      title={`关于 ${APP_NAME}`}
      onClose={close}
      width={480}
      footer={
        <>
          <Button variant="ghost" onClick={() => void openReleasesPage(PROJECT_URL)}>
            项目主页
          </Button>
          <Button
            onClick={() => {
              void useUpdateStore.getState().check();
            }}
          >
            {checking ? "检查中…" : "检查更新"}
          </Button>
          <Button variant="primary" autoFocus onClick={close}>
            关闭
          </Button>
        </>
      }
    >
      <div className="relative">
        {confetti.length > 0 ? (
          <div className="pointer-events-none absolute inset-0 z-10 overflow-hidden">
            {confetti.map((piece) => (
              <span
                key={piece.id}
                className="mm-confetti-piece"
                style={{
                  left: `${piece.left}%`,
                  width: piece.width,
                  height: piece.height,
                  background: piece.color,
                  animationDelay: `${piece.delay}s`,
                  animationDuration: `${piece.duration}s`,
                  transform: `rotate(${piece.rotate}deg)`,
                }}
              />
            ))}
          </div>
        ) : null}

        <div className="space-y-4">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={handleLogoClick}
              title="点击试试？"
              className="flex h-12 w-12 shrink-0 items-center justify-center rounded-[12px] transition-transform hover:scale-105 active:scale-95"
            >
              <img src="/app-icon.png" alt={APP_NAME} className="h-12 w-12 rounded-[12px]" />
            </button>
            <div className="min-w-0 flex-1">
              <div className="flex items-center gap-2">
                <span className="text-[16px] font-semibold text-fg">{APP_NAME}</span>
                {eggUnlocked ? (
                  <span className="rounded bg-accent-soft px-1.5 py-0.5 text-[10px] text-accent">
                    ✨ 彩蛋已解锁
                  </span>
                ) : null}
              </div>
              <div className="text-[12px] text-muted">Windows 全能文档查看与编辑器（Markdown / PDF / 代码 / 表格）</div>
              <div className="mt-0.5 font-mono text-[11px] text-faint">
                版本 {version || "…"}
                {info?.hasUpdate ? ` · 有新版本 ${info.latest}` : ""}
              </div>
            </div>
            {hints ? <span className="shrink-0 text-[11px] text-faint">{hints}</span> : null}
          </div>

          <div className="rounded-[var(--radius)] border border-line bg-panel px-3 py-2.5 text-[12px]">
            <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-faint">
              作者
            </div>
            <div className="text-fg">
              {AUTHOR_NAME}（{AUTHOR_NAME_CN}）
            </div>
            <div className="mt-0.5 text-[11px] text-muted">
              Master 系列软件作者 · 本软件由其设计与开发
            </div>
          </div>

          <div className="rounded-[var(--radius)] border border-line bg-panel px-3 py-2.5 text-[12px]">
            <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-faint">
              软件中心
            </div>
            <button
              type="button"
              onClick={() => void openReleasesPage(PORTAL_URL)}
              className="text-left text-accent hover:underline"
            >
              {PORTAL_URL}
            </button>
            <div className="mt-0.5 text-[11px] text-muted">
              Master 系列软件统一下载与更新入口（含国内加速节点，应用内更新即从此获取）
            </div>
          </div>

          <div className="rounded-[var(--radius)] border border-line bg-panel px-3 py-2.5 text-[12px]">
            <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-faint">
              便携版
            </div>
            <div className="text-muted">
              单文件便携版：双击即用，无需安装与卸载，可放在任意目录（含 U 盘）。
              应用内更新会自动替换当前程序并重启。
            </div>
            {exePath ? (
              <div className="mt-1 break-all font-mono text-[10px] text-faint">{exePath}</div>
            ) : null}
          </div>

          {eggUnlocked ? (
            <div className="rounded-[var(--radius)] border border-accent bg-accent-soft px-3 py-3 text-[12px]">
              <div className="mb-1 flex items-center gap-1.5 font-semibold text-accent">
                🎉 彩蛋解锁 · 作者的话
              </div>
              <div className="text-fg">
                「把复杂留给自己，把简单留给用户。」
              </div>
              <div className="mt-1 text-[11px] text-muted">
                —— {AUTHOR_NAME}（{AUTHOR_NAME_CN}）· 2026
              </div>
              <div className="mt-2 text-[11px] text-muted">
                {APP_NAME} 是 Master 系列软件的一员，愿你写得顺手、读得舒心。
              </div>
              <div className="mt-2 flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={() => {
                    useDungeonStore.getState().open();
                    close();
                  }}
                  className="rounded-md border border-transparent bg-accent px-2.5 py-1 text-[11px] font-semibold text-accent-fg transition-opacity hover:opacity-90"
                  >
                    🗝️ Markdown 地牢（小游戏）
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      useAppStore.getState().openGame();
                      void import("../../wuxia/store").then((m) => m.useWuxiaStore.getState().ensureLoaded());
                      close();
                    }}
                    className="rounded-md border border-line bg-elevated px-2.5 py-1 text-[11px] font-semibold text-fg hover:bg-hover"
                  >
                    🏯 Markdown 江湖（文字武侠）
                  </button>
                <button
                  type="button"
                  onClick={() => setConfetti(makeConfetti())}
                  className="rounded-md border border-line bg-elevated px-2 py-1 text-[11px] text-fg hover:bg-hover"
                >
                  再撒一次彩纸 🎊
                </button>
                {dungeonMeta.cleared ? (
                  <span className="rounded bg-elevated px-1.5 py-0.5 text-[10px] text-fg">
                    🏆 地牢已通关
                  </span>
                ) : null}
              </div>
            </div>
          ) : (
            <div className="space-y-1 text-[11px] leading-relaxed text-muted">
              <div>
                技术栈：Tauri 2 · Rust · React 19 · TypeScript · CodeMirror 6 ·
                markdown-it
              </div>
              <div>
                项目地址：
                <button
                  type="button"
                  className="ml-1 text-accent hover:underline"
                  onClick={() => void openReleasesPage(PROJECT_URL)}
                >
                  {PROJECT_URL.replace("https://", "")}
                </button>
              </div>
              <div>© 2026 {AUTHOR_NAME}（{AUTHOR_NAME_CN}）· 保留所有权利</div>
            </div>
          )}
        </div>
      </div>
    </Modal>
  );
}
