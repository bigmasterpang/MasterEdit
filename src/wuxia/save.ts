/**
 * 《Markdown 江湖》存档：localStorage + 版本迁移
 */

import { createPlayer, emptyIdle } from "./engine";
import { FIRST_MAIN_QUEST, MAP_BY_ID } from "./data";
import { readMigratedItem, removeItem, writeItem } from "../utils/storage";
import type { SaveGame } from "./types";

const STORAGE_KEY = "masteredit.wuxia.save.v1";
/** 改名前的键名，读取时兼容迁移 */
const LEGACY_STORAGE_KEY = "mastermd.wuxia.save.v1";
export const SAVE_VERSION = 1;

export function todayKey(): string {
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

export function newGame(playerName = "无名少侠"): SaveGame {
  const player = createPlayer(playerName);
  const now = Date.now();
  return {
    version: SAVE_VERSION,
    createdAt: now,
    lastSeen: now,
    player,
    quests: { active: [FIRST_MAIN_QUEST], progress: {}, completed: [], repeatUsed: {} },
    maps: [MAP_BY_ID.qingshi ? "qingshi" : ""].filter(Boolean),
    autoAccept: true,
    idle: emptyIdle("qingshi"),
    dungeonRun: null,
    dungeonDaily: {},
    stats: { kills: 0, deaths: 0, dungeonClears: 0, battles: 0, playMs: 0 },
    achievements: [],
    title: "江湖新兵",
  };
}

export function loadGame(): SaveGame | null {
  try {
    const raw = readMigratedItem(STORAGE_KEY, LEGACY_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Partial<SaveGame>;
    if (!parsed || !parsed.player) return null;
    return migrate(parsed);
  } catch (error) {
    console.error("[江湖] 读取存档失败", error);
    return null;
  }
}

export function saveGame(save: SaveGame): void {
  try {
    writeItem(STORAGE_KEY, JSON.stringify({ ...save, lastSeen: Date.now() }));
  } catch (error) {
    console.error("[江湖] 写入存档失败", error);
  }
}

export function clearSave(): void {
  removeItem(STORAGE_KEY, LEGACY_STORAGE_KEY);
}

/** 版本迁移：字段缺失时补默认值，保证旧存档可用 */
function migrate(input: Partial<SaveGame>): SaveGame {
  const base = newGame(input.player?.name ?? "无名少侠");
  const save: SaveGame = {
    ...base,
    ...input,
    version: SAVE_VERSION,
    player: { ...base.player, ...input.player },
    quests: { ...base.quests, ...input.quests, repeatUsed: input.quests?.repeatUsed ?? {} },
    idle: {
      ...base.idle,
      ...input.idle,
      config: { ...base.idle.config, ...input.idle?.config },
      session: { ...base.idle.session, ...input.idle?.session },
    },
    stats: { ...base.stats, ...input.stats },
    dungeonDaily: input.dungeonDaily ?? {},
    maps: input.maps && input.maps.length > 0 ? input.maps : base.maps,
    autoAccept: input.autoAccept ?? base.autoAccept,
    achievements: input.achievements ?? [],
    dungeonRun: input.dungeonRun ?? null,
    title: input.title ?? "江湖新兵",
  };
  return save;
}
