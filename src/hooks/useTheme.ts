import { useEffect, useState } from "react";
import { useSettingsStore } from "../stores/settingsStore";

/** 解析主题模式为实际是否深色，并应用到 <html> */
export function useTheme(): boolean {
  const theme = useSettingsStore((s) => s.theme);
  const accent = useSettingsStore((s) => s.accent);
  const fontSize = useSettingsStore((s) => s.fontSize);
  const fontFamily = useSettingsStore((s) => s.fontFamily);
  const colorblindMode = useSettingsStore((s) => s.colorblindMode);
  const [systemDark, setSystemDark] = useState(() =>
    typeof window !== "undefined"
      ? window.matchMedia("(prefers-color-scheme: dark)").matches
      : false,
  );

  useEffect(() => {
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    const onChange = (e: MediaQueryListEvent) => setSystemDark(e.matches);
    mq.addEventListener("change", onChange);
    return () => mq.removeEventListener("change", onChange);
  }, []);

  const isDark = theme === "dark" || (theme === "system" && systemDark);

  useEffect(() => {
    const root = document.documentElement;
    root.classList.toggle("dark", isDark);
    root.dataset.accent = accent;
    root.dataset.colorblind = colorblindMode ? "true" : "false";
    root.style.setProperty("--editor-size", `${fontSize}px`);
    if (fontFamily.trim()) {
      root.style.setProperty("--font-mono", fontFamily.trim());
    } else {
      root.style.removeProperty("--font-mono");
    }
  }, [isDark, accent, fontSize, fontFamily, colorblindMode]);

  return isDark;
}

