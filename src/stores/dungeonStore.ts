import { create } from "zustand";
import {
  BOSS_FLOOR,
  choose,
  finishRun,
  newRun,
  type DungeonState,
} from "../easter-egg/dungeon";
import { readMigratedItem, writeItem } from "../utils/storage";

const STORAGE_KEY = "masteredit.dungeon.v1";
/** 改名前的键名，读取时兼容迁移 */
const LEGACY_STORAGE_KEY = "mastermd.dungeon.v1";

interface DungeonMeta {
  runs: number;
  cleared: boolean;
  bestFloor: number;
}

function loadMeta(): DungeonMeta {
  try {
    const raw = readMigratedItem(STORAGE_KEY, LEGACY_STORAGE_KEY);
    if (!raw) return { runs: 0, cleared: false, bestFloor: 0 };
    const parsed = JSON.parse(raw) as Partial<DungeonMeta>;
    return {
      runs: Number(parsed.runs ?? 0),
      cleared: Boolean(parsed.cleared),
      bestFloor: Number(parsed.bestFloor ?? 0),
    };
  } catch {
    return { runs: 0, cleared: false, bestFloor: 0 };
  }
}

function saveMeta(meta: DungeonMeta): void {
  writeItem(STORAGE_KEY, JSON.stringify(meta));
}

interface DungeonStore {
  visible: boolean;
  state: DungeonState;
  meta: DungeonMeta;

  open: () => void;
  close: () => void;
  startRun: () => void;
  act: (action: string) => void;
}

export const useDungeonStore = create<DungeonStore>((set, get) => ({
  visible: false,
  state: newRun(),
  meta: loadMeta(),

  open: () => set({ visible: true }),
  close: () => set({ visible: false }),

  startRun: () => {
    const meta = get().meta;
    const next = { ...meta, runs: meta.runs + 1 };
    saveMeta(next);
    set({ state: newRun(), meta: next });
  },

  act: (action) => {
    const current = get().state;
    if (action === "restart") {
      get().startRun();
      return;
    }
    let next = choose(current, action);
    // 击败 Boss 后进入结算
    if (next.floor > BOSS_FLOOR && next.phase !== "victory" && next.phase !== "gameover") {
      next = finishRun(next);
    }
    if (next.phase === "victory") {
      const meta = get().meta;
      if (!meta.cleared) {
        const updated = { ...meta, cleared: true, bestFloor: BOSS_FLOOR };
        saveMeta(updated);
        set({ state: next, meta: updated });
        return;
      }
    }
    if (next.phase === "gameover") {
      const meta = get().meta;
      if (next.floor > meta.bestFloor) {
        const updated = { ...meta, bestFloor: next.floor };
        saveMeta(updated);
        set({ state: next, meta: updated });
        return;
      }
    }
    set({ state: next });
  },
}));
