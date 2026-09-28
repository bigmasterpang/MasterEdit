import { useEffect, useState } from "react";
import { Modal, Button } from "../common/Modal";
import { Icon } from "../common/Icon";
import { useDialogStore } from "../../stores/dialogStore";
import { useSettingsStore } from "../../stores/settingsStore";
import { useUpdateStore } from "../../stores/updateStore";
import type { AccentName, ThemeMode } from "../../types";
import {
  APP_NAME,
  AUTHOR_NAME,
  AUTHOR_NAME_CN,
  MAX_FONT_SIZE,
  MIN_FONT_SIZE,
  MIN_AUTOSAVE_INTERVAL,
} from "../../utils/constants";
import { getVersion } from "@tauri-apps/api/app";

const THEMES: Array<{ value: ThemeMode; label: string }> = [
  { value: "light", label: "浅色" },
  { value: "dark", label: "深色" },
  { value: "system", label: "跟随系统" },
];

const ACCENTS: Array<{
  value: AccentName;
  label: string;
  cbLabel: string;
  color: string;
  cbColor: string;
}> = [
  { value: "blue", label: "蓝", cbLabel: "钴蓝", color: "#2563eb", cbColor: "#2563eb" },
  { value: "violet", label: "紫", cbLabel: "靛紫", color: "#7c3aed", cbColor: "#7c3aed" },
  { value: "emerald", label: "绿", cbLabel: "天青", color: "#059669", cbColor: "#0284c7" },
  { value: "amber", label: "橙", cbLabel: "琥珀", color: "#d97706", cbColor: "#d97706" },
  { value: "rose", label: "红", cbLabel: "洋红", color: "#e11d48", cbColor: "#c026d3" },
];

const FONT_PRESETS = [
  { label: "默认", value: "" },
  { label: "Cascadia Mono", value: '"Cascadia Mono", Consolas, monospace' },
  { label: "Consolas", value: "Consolas, monospace" },
  { label: "JetBrains Mono", value: '"JetBrains Mono", Consolas, monospace' },
  { label: "微软雅黑", value: '"Microsoft YaHei", sans-serif' },
];

export function SettingsModal() {
  const open = useDialogStore((s) => s.settingsVisible);
  const close = () => useDialogStore.getState().setSettingsVisible(false);
  const settings = useSettingsStore();
  const updateState = useUpdateStore();
  const [version, setVersion] = useState("");

  useEffect(() => {
    void getVersion()
      .then(setVersion)
      .catch(() => setVersion("0.1.0"));
  }, []);

  const update = (patch: Parameters<typeof settings.update>[0]) => {
    settings.update(patch);
  };

  return (
    <Modal
      open={open}
      title="设置"
      onClose={close}
      width={560}
      footer={
        <>
          <Button
            variant="ghost"
            onClick={() => {
              settings.reset();
            }}
          >
            恢复默认
          </Button>
          <Button variant="primary" autoFocus onClick={close}>
            完成
          </Button>
        </>
      }
    >
      <div className="space-y-5">
        <Section title="外观与无障碍">
          <Row label="主题">
            <Segmented
              items={THEMES}
              value={settings.theme}
              onChange={(value) => update({ theme: value })}
            />
          </Row>
          <Row label="配色方案">
            <div className="flex flex-wrap items-center gap-1.5">
              {ACCENTS.map((accent) => {
                const active = settings.accent === accent.value;
                const displayColor = settings.colorblindMode ? accent.cbColor : accent.color;
                const displayLabel = settings.colorblindMode ? accent.cbLabel : accent.label;
                return (
                  <button
                    key={accent.value}
                    type="button"
                    title={displayLabel}
                    onClick={() => update({ accent: accent.value })}
                    className={`flex items-center gap-1.5 rounded-full border px-2 py-0.5 text-[11.5px] transition-all ${
                      active
                        ? "border-accent bg-accent-soft font-medium text-fg"
                        : "border-line bg-input text-muted hover:border-line-strong hover:text-fg"
                    }`}
                  >
                    <span
                      className="h-3 w-3 shrink-0 rounded-full border border-black/10"
                      style={{ background: displayColor }}
                    />
                    <span>{displayLabel}</span>
                  </button>
                );
              })}
            </div>
          </Row>
          <div className="rounded-lg border border-line bg-panel/60 px-3 py-2">
            <div className="flex items-center justify-between gap-4">
              <div className="min-w-0">
                <div className="flex items-center gap-1.5 text-[12px] font-medium text-fg">
                  <Icon name="eye" size={13} className="text-accent" />
                  <span>色弱友好模式（红绿色弱优化）</span>
                </div>
                <div className="mt-0.5 text-[11px] leading-relaxed text-faint">
                  启用 Okabe-Ito 无障碍高辨识度色板（以天青蓝、琥珀金、洋红紫替代易混淆的红绿对比），并为代码语法高亮、提示块及 PDF 批注增加字重与线条样式辅助标识。
                </div>
              </div>
              <Toggle
                checked={settings.colorblindMode}
                onChange={(value) => update({ colorblindMode: value })}
              />
            </div>
          </div>
          <Row label={`字号（${settings.fontSize}px）`}>
            <input
              type="range"
              min={MIN_FONT_SIZE}
              max={MAX_FONT_SIZE}
              value={settings.fontSize}
              onChange={(event) => update({ fontSize: Number(event.target.value) })}
              className="w-40 accent-[var(--accent)]"
            />
          </Row>
          <Row label="编辑区字体">
            <select
              value={settings.fontFamily}
              onChange={(event) => update({ fontFamily: event.target.value })}
              className="w-52 rounded-md border border-line bg-input px-2 py-1 text-[12px] text-fg"
            >
              {FONT_PRESETS.map((preset) => (
                <option key={preset.label} value={preset.value}>
                  {preset.label}
                </option>
              ))}
            </select>
          </Row>
        </Section>

        <Section title="编辑器">
          <Row label={`缩进宽度（${settings.tabSize} 空格）`}>
            <input
              type="range"
              min={1}
              max={8}
              value={settings.tabSize}
              onChange={(event) => update({ tabSize: Number(event.target.value) })}
              className="w-40 accent-[var(--accent)]"
            />
          </Row>
          <Row label="自动换行">
            <Toggle
              checked={settings.wordWrap}
              onChange={(value) => update({ wordWrap: value })}
            />
          </Row>
          <Row label="显示行号">
            <Toggle
              checked={settings.showLineNumbers}
              onChange={(value) => update({ showLineNumbers: value })}
            />
          </Row>
        </Section>

        <Section title="保存">
          <Row label="自动保存">
            <Toggle
              checked={settings.autoSave}
              onChange={(value) => update({ autoSave: value })}
            />
          </Row>
          <Row label={`自动保存间隔（${settings.autoSaveInterval} 秒）`}>
            <input
              type="range"
              min={MIN_AUTOSAVE_INTERVAL}
              max={300}
              step={5}
              value={settings.autoSaveInterval}
              disabled={!settings.autoSave}
              onChange={(event) =>
                update({ autoSaveInterval: Number(event.target.value) })
              }
              className="w-40 accent-[var(--accent)] disabled:opacity-40"
            />
          </Row>
          <Row label={`最近文件上限（${settings.recentFilesLimit}）`}>
            <input
              type="range"
              min={1}
              max={10}
              value={settings.recentFilesLimit}
              onChange={(event) =>
                update({ recentFilesLimit: Number(event.target.value) })
              }
              className="w-40 accent-[var(--accent)]"
            />
          </Row>
        </Section>

        <Section title="快捷键">
          <Row label="可自定义绑定与全部快捷键参考">
            <button
              type="button"
              onClick={() => useDialogStore.getState().setShortcutsVisible(true)}
              className="flex items-center gap-1.5 rounded-md border border-line bg-input px-2.5 py-1 text-[12px] text-fg hover:bg-hover"
            >
              <Icon name="keyboard" size={13} />
              打开快捷键面板
            </button>
          </Row>
        </Section>

        <Section title="版本更新">
          <Row label="启动时自动检查更新">
            <Toggle
              checked={settings.autoCheckUpdate}
              onChange={(value) => update({ autoCheckUpdate: value })}
            />
          </Row>
          <Row label={`当前版本 ${version}`}>
            <button
              type="button"
              onClick={() => void useUpdateStore.getState().check()}
              className="flex items-center gap-1.5 rounded-md border border-line bg-input px-2.5 py-1 text-[12px] text-fg hover:bg-hover disabled:opacity-50"
              disabled={updateState.checking}
            >
              <Icon name={updateState.checking ? "loader" : "refresh"} size={13} />
              {updateState.checking ? "检查中…" : "检查更新"}
            </button>
          </Row>
          <div className="text-[11px] text-faint">
            版本来源：软件中心 · github.com/bigmasterpang/MasterEdit
            {updateState.info?.hasUpdate
              ? ` · 发现新版本 ${updateState.info.latest}`
              : updateState.info
                ? " · 已是最新版本"
                : ""}
          </div>
        </Section>

        <Section title="关于">
          <Row label={`${APP_NAME} · 作者 ${AUTHOR_NAME}（${AUTHOR_NAME_CN}）`}>
            <button
              type="button"
              onClick={() => useDialogStore.getState().setAboutVisible(true)}
              className="flex items-center gap-1.5 rounded-md border border-line bg-input px-2.5 py-1 text-[12px] text-fg hover:bg-hover"
            >
              <Icon name="info" size={13} />
              关于 {APP_NAME}
            </button>
          </Row>
        </Section>
      </div>
    </Modal>
  );
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section>
      <div className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-faint">
        {title}
      </div>
      <div className="space-y-2">{children}</div>
    </section>
  );
}

function Row({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex items-center justify-between gap-4">
      <span className="text-[12px] text-muted">{label}</span>
      {children}
    </div>
  );
}

function Segmented<T extends string>({
  items,
  value,
  onChange,
}: {
  items: Array<{ value: T; label: string }>;
  value: T;
  onChange: (value: T) => void;
}) {
  return (
    <div className="flex rounded-md border border-line bg-input p-0.5">
      {items.map((item) => (
        <button
          key={item.value}
          type="button"
          onClick={() => onChange(item.value)}
          className={`rounded px-2.5 py-1 text-[12px] transition-colors ${
            value === item.value
              ? "bg-accent-soft-strong font-medium text-accent"
              : "text-muted hover:text-fg"
          }`}
        >
          {item.label}
        </button>
      ))}
    </div>
  );
}

function Toggle({
  checked,
  onChange,
}: {
  checked: boolean;
  onChange: (value: boolean) => void;
}) {
  return (
    <button
      type="button"
      role="switch"
      aria-checked={checked}
      onClick={() => onChange(!checked)}
      className={`relative h-5 w-9 shrink-0 rounded-full transition-colors ${
        checked ? "bg-accent" : "bg-line-strong"
      }`}
    >
      <span
        className={`absolute top-0.5 h-4 w-4 rounded-full bg-white transition-all ${
          checked ? "left-[18px]" : "left-0.5"
        }`}
      />
    </button>
  );
}
