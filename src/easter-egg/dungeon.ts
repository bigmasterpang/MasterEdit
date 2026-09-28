/**
 * 彩蛋：Markdown 地牢（文字冒险）
 *
 * 整个玩法都用 Markdown 语法演出：引用块做旁白、任务列表做背包、表格做属性、
 * 代码块做战斗日志、提示块做警告，正好展示本软件的渲染能力。
 *
 * 引擎为纯函数：给定当前状态与选择，返回新状态；UI 只负责渲染。
 */

export interface Shard {
  id: string;
  name: string;
  markdown: string;
  desc: string;
}

/** 语法碎片（技能）：收集后可在地牢中使用 */
export const SHARDS: Shard[] = [
  { id: "bold", name: "粗体之力", markdown: "**粗体**", desc: "永久攻击 +2" },
  { id: "italic", name: "斜体之疾", markdown: "*斜体*", desc: "使用后必定闪避下一次攻击" },
  { id: "strike", name: "删除之刃", markdown: "~~删除线~~", desc: "技能攻击无视护甲" },
  { id: "code", name: "代码之盾", markdown: "`行内代码`", desc: "每场战斗首次受击免伤" },
  { id: "sup", name: "上标之速", markdown: "^上标^", desc: "技能攻击有 50% 概率追击" },
  { id: "mark", name: "高亮之眼", markdown: "==高亮==", desc: "看穿弱点，对弱点怪物伤害翻倍" },
];

export interface Monster {
  name: string;
  hp: number;
  atk: number;
  def: number;
  /** 弱点碎片 id：用对应技能攻击可造成额外伤害 */
  weak: string;
  taunt: string;
}

const MONSTERS: Monster[] = [
  {
    name: "断链蜘蛛",
    hp: 14,
    atk: 3,
    def: 0,
    weak: "strike",
    taunt: "它用残缺的 `](` 织网，链接一个也点不开。",
  },
  {
    name: "乱码史莱姆",
    hp: 18,
    atk: 4,
    def: 1,
    weak: "code",
    taunt: "它全身都是 \uFFFD，碰一下你的编码就乱。",
  },
  {
    name: "Tab 幽灵",
    hp: 22,
    atk: 5,
    def: 1,
    weak: "italic",
    taunt: "它把缩进搅成空格与制表符的战场。",
  },
  {
    name: "缺行怪兽",
    hp: 26,
    atk: 6,
    def: 2,
    weak: "sup",
    taunt: "它删掉了你的空行，让段落挤成一团。",
  },
];

const BOSS: Monster = {
  name: "格式混乱之兽",
  hp: 46,
  atk: 7,
  def: 2,
  weak: "mark",
  taunt: "粗体与斜体纠缠，表格歪斜，标题层级全乱——文档的噩梦本体。",
};

export type Phase = "intro" | "explore" | "combat" | "reward" | "gameover" | "victory";

export interface DungeonState {
  floor: number;
  hp: number;
  maxHp: number;
  atk: number;
  ink: number;
  shards: string[];
  shield: boolean;
  dodge: boolean;
  monster: Monster | null;
  monsterHp: number;
  phase: Phase;
  log: string[];
  scene: string;
  choices: Choice[];
  /** 通关/失败后的总结 */
  result: string;
  /** 本局获得的碎片（用于结算展示） */
  gained: string[];
  lastBookLine: string;
}

export interface Choice {
  label: string;
  hint?: string;
  /** 引擎内部使用 */
  action: string;
}

/* ------------------------------ 工具函数 ------------------------------ */

const rand = (min: number, max: number) => Math.floor(Math.random() * (max - min + 1)) + min;

const pick = <T,>(list: T[]): T => list[rand(0, list.length - 1)];

const shardOf = (id: string) => SHARDS.find((s) => s.id === id);

const BOOK_LINES = [
  "「把复杂留给自己，把简单留给用户。」",
  "「好的文档，是写给三个月后的自己。」",
  "「Markdown 的魅力在于：你只管写，格式交给渲染。」",
  "「删掉一个 ~~段落~~，往往比加一个更勇敢。」",
  "「`Ctrl+S` 是最廉价的后悔药。」",
  "「别人看的是结果，你自己知道那些 `==高亮==` 意味着什么。」",
];

/* ------------------------------ Markdown 组装 ------------------------------ */

function shardListMarkdown(state: DungeonState): string {
  if (state.shards.length === 0) return "- [ ] 尚未收集任何语法碎片";
  return state.shards
    .map((id) => {
      const shard = shardOf(id);
      return `- [x] ${shard?.markdown ?? id} · ${shard?.name ?? ""}`;
    })
    .join("\n");
}

function statTable(state: DungeonState): string {
  const shown = state.monster
    ? `| 你 | ${state.hp}/${state.maxHp} | ${state.atk} | ${state.ink} |\n| ${state.monster.name} | ${state.monsterHp} | ${state.monster.atk} | ${state.monster.def} |`
    : `| 你 | ${state.hp}/${state.maxHp} | ${state.atk} | ${state.ink} |`;
  return [
    "| 角色 | HP | 攻击 | 墨水 |",
    "| --- | --- | --- | --- |",
    shown,
  ].join("\n");
}

function logBlock(state: DungeonState): string {
  if (state.log.length === 0) return "";
  return ["```text", ...state.log.slice(0, 8), "```"].join("\n");
}

function combatMarkdown(state: DungeonState): string {
  const monster = state.monster;
  if (!monster) return "";
  const weakShard = shardOf(monster.weak);
  const knowWeakness = state.shards.includes("mark");
  const weakText = knowWeakness
    ? `\n> [!WARNING] 高亮之眼洞察到它的弱点：${weakShard?.markdown ?? ""}（用对应技能伤害翻倍）`
    : "";
  return [
    `### 第 ${state.floor} 层 · 遭遇 ${monster.name}`,
    "",
    `> ${monster.taunt}`,
    weakText,
    "",
    statTable(state),
    "",
    "**背包**",
    "",
    shardListMarkdown(state),
    "",
    logBlock(state),
  ]
    .filter((line) => line !== undefined)
    .join("\n");
}

function exploreMarkdown(state: DungeonState, title: string, body: string): string {
  return [
    `### 第 ${state.floor} 层 · ${title}`,
    "",
    body,
    "",
    statTable(state),
    "",
    "**背包**",
    "",
    shardListMarkdown(state),
    "",
    logBlock(state),
  ].join("\n");
}

/* ------------------------------ 事件构造 ------------------------------ */

function makeCombat(state: DungeonState, monster: Monster, intro: string): DungeonState {
  const next: DungeonState = {
    ...state,
    phase: "combat",
    monster,
    monsterHp: monster.hp,
    shield: state.shards.includes("code"),
    dodge: false,
    log: [intro],
  };
  return {
    ...next,
    scene: combatMarkdown(next),
    choices: combatChoices(next),
  };
}

function combatChoices(state: DungeonState): Choice[] {
  const choices: Choice[] = [{ label: "挥剑攻击", hint: "普通攻击", action: "attack" }];
  const usable = state.shards.filter((id) => id !== "bold" && id !== "code");
  for (const id of usable) {
    const shard = shardOf(id);
    if (!shard) continue;
    choices.push({
      label: `释放 ${shard.markdown} 技能`,
      hint: shard.desc,
      action: `skill:${id}`,
    });
  }
  choices.push({ label: `喝一瓶墨水（剩 ${state.ink}）`, hint: "恢复 10 HP", action: "drink" });
  choices.push({ label: "逃跑", hint: "50% 成功，失败会挨打", action: "flee" });
  return choices;
}

function drawEvent(state: DungeonState): DungeonState {
  if (state.floor >= 5) {
    return makeCombat(state, BOSS, "地牢尽头，一团扭曲的排版正在蠕动。");
  }
  const roll = rand(1, 100);
  if (roll <= 55) {
    const monster = pick(MONSTERS.filter((m) => m.hp >= state.floor * 10));
    return makeCombat(state, monster ?? MONSTERS[0], "阴影里传来键盘敲击声。");
  }
  if (roll <= 72) return openChest(state);
  if (roll <= 88) return restSite(state);
  return bookshelf(state);
}

function openChest(state: DungeonState): DungeonState {
  const next: DungeonState = {
    ...state,
    phase: "explore",
    monster: null,
    log: ["你发现一只被 ``` 封住的箱子。"],
    choices: [
      { label: "打开箱子", action: "chest:open" },
      { label: "谨慎离开", hint: "避免陷阱", action: "chest:leave" },
    ],
  };
  return {
    ...next,
    scene: exploreMarkdown(
      next,
      "密封的箱子",
      "> 箱盖上写着：`<!-- 危险：内含未转义的内容 -->`\n\n打开它？",
    ),
  };
}

function restSite(state: DungeonState): DungeonState {
  const next: DungeonState = {
    ...state,
    phase: "explore",
    monster: null,
    log: ["篝火在黑暗中噼啪作响。"],
    choices: [
      { label: "休息（回满 HP）", action: "rest:sleep" },
      { label: "磨剑（攻击 +1）", action: "rest:sharpen" },
    ],
  };
  return {
    ...next,
    scene: exploreMarkdown(
      next,
      "篝火",
      "> 火焰的形状像一段排版工整的列表。\n\n你要休息，还是磨快手中的武器？",
    ),
  };
}

function bookshelf(state: DungeonState): DungeonState {
  const line = pick(BOOK_LINES);
  const next: DungeonState = {
    ...state,
    phase: "explore",
    monster: null,
    lastBookLine: line,
    log: ["书架上全是关于写作的书。"],
    choices: [
      { label: "翻阅（恢复 5 HP）", action: "book:read" },
      { label: "离开", action: "book:leave" },
    ],
  };
  return {
    ...next,
    scene: exploreMarkdown(
      next,
      "神秘书架",
      `> 一本无名之书自动翻开，页面上写着一行字：\n>\n> ${line}`,
    ),
  };
}

/* ------------------------------ 战斗结算 ------------------------------ */

function damageToMonster(state: DungeonState, base: number, ignoreDef: boolean): number {
  const def = ignoreDef ? 0 : (state.monster?.def ?? 0);
  return Math.max(1, base - def);
}

function monsterTurn(state: DungeonState, log: string[]): { state: DungeonState; log: string[] } {
  const monster = state.monster;
  if (!monster) return { state, log };
  if (state.dodge) {
    log.push(`你以 *斜体* 的身法闪过 ${monster.name} 的攻击。`);
    return { state: { ...state, dodge: false }, log };
  }
  if (state.shield) {
    log.push("`代码之盾` 挡下了第一击，护盾碎裂。");
    return { state: { ...state, shield: false }, log };
  }
  const damage = Math.max(1, monster.atk + rand(0, 2) - rand(0, 1));
  const hp = state.hp - damage;
  log.push(`${monster.name} 反击，造成 ${damage} 点伤害（HP ${Math.max(0, hp)}）。`);
  return { state: { ...state, hp }, log };
}

function afterCombat(state: DungeonState): DungeonState {
  const remain = SHARDS.filter((s) => !state.shards.includes(s.id) && s.id !== "bold");
  const rewardPool = remain.length > 0 ? remain : SHARDS.filter((s) => s.id === "bold");
  const reward = pick(rewardPool);
  const alreadyHave = state.shards.includes(reward.id);
  const nextShards = alreadyHave ? state.shards : [...state.shards, reward.id];
  const atkBonus = reward.id === "bold" && !alreadyHave ? 2 : 0;
  const inkGain = alreadyHave ? 2 : 1;

  const next: DungeonState = {
    ...state,
    phase: "reward",
    monster: null,
    monsterHp: 0,
    shards: nextShards,
    atk: state.atk + atkBonus,
    ink: state.ink + inkGain,
    gained: alreadyHave ? state.gained : [...state.gained, reward.id],
    log: [
      ...state.log,
      `你击败了敌人。战利品：${reward.markdown}（${reward.name}）${atkBonus ? "，攻击 +2" : ""}，墨水 +${inkGain}`,
    ],
    choices: [{ label: "继续深入", action: "descend" }],
  };
  return {
    ...next,
    scene: exploreMarkdown(
      next,
      "战斗胜利",
      `> 你捡起了散落在地上的语法碎片：${reward.markdown}\n>\n> ${reward.desc}\n\n${
        state.shards.includes("mark") ? "高亮之眼提示：更深处的敌人更强。" : ""
      }`,
    ),
  };
}

/* ------------------------------ 对外 API ------------------------------ */

export function newRun(): DungeonState {
  const base: DungeonState = {
    floor: 1,
    hp: 30,
    maxHp: 30,
    atk: 5,
    ink: 2,
    shards: [],
    shield: false,
    dodge: false,
    monster: null,
    monsterHp: 0,
    phase: "intro",
    log: [],
    scene: "",
    choices: [],
    result: "",
    gained: [],
    lastBookLine: "",
  };
  return {
    ...base,
    scene: [
      "# 🗝️ Markdown 地牢",
      "",
      "> 你在一份被遗忘的文档深处醒来，四周全是错乱的排版。",
      "> 传说走到第 **5** 层，就能见到「==文档之心==」，",
      "> 但路上有各种因格式而生的怪物。",
      "",
      "**规则**",
      "",
      "1. 每层遭遇战斗 / 宝箱 / 篝火 / 书架之一",
      "2. 战斗胜利可获得语法碎片（技能）",
      "3. 收集到 `==高亮==` 后可以看穿怪物弱点，伤害翻倍",
      "4. 血量归零则本局结束，随时可以重来",
      "",
      "| 你的初始属性 | 数值 |",
      "| --- | --- |",
      "| HP | 30 |",
      "| 攻击 | 5 |",
      "| 墨水（治疗） | 2 |",
    ].join("\n"),
    choices: [{ label: "踏入门内", action: "start" }],
  };
}

export function choose(state: DungeonState, action: string): DungeonState {
  if (action === "start") return drawEvent({ ...state, log: ["你推开门，走了进去。"] });
  if (action === "descend") {
    return drawEvent({ ...state, floor: state.floor + 1 });
  }
  if (action === "restart") return newRun();

  // ---- 战斗动作 ----
  if (state.phase === "combat" && state.monster) {
    const monster = state.monster;
    let next = { ...state };
    const log = [...state.log];

    if (action === "drink") {
      if (state.ink <= 0) {
        log.push("墨水用完了，你只能握紧武器。");
      } else {
        const healed = Math.min(state.maxHp, state.hp + 10);
        log.push(`你喝下墨水，恢复了 ${healed - state.hp} 点 HP。`);
        next = { ...next, hp: healed, ink: state.ink - 1 };
      }
    } else if (action === "flee") {
      if (rand(1, 100) <= 50) {
        log.push("你成功逃走了，但宝物也与你无关。");
        const escaped: DungeonState = {
          ...state,
          floor: state.floor, // 逃跑不推进层数
          monster: null,
          monsterHp: 0,
          phase: "reward",
          log,
          choices: [{ label: "继续深入", action: "descend" }],
        };
        return {
          ...escaped,
          scene: exploreMarkdown(escaped, "逃出生天", "> 你气喘吁吁地靠在墙边。\n\n虽然没有战利品，但你还活着。"),
        };
      }
      log.push("逃跑失败，怪物挡住了退路！");
    } else if (action.startsWith("skill:")) {
      const shardId = action.slice("skill:".length);
      const isWeak = shardId === monster.weak;
      const insight = state.shards.includes("mark") && isWeak;
      const ignoreDef = shardId === "strike";
      const base = state.atk + 3 + rand(0, 2) + (insight ? 3 : 0);
      let damage = damageToMonster(state, base, ignoreDef);
      if (insight) damage *= 2;
      next = { ...next, monsterHp: next.monsterHp - damage };
      const shard = shardOf(shardId);
      log.push(
        `你释放 ${shard?.markdown ?? shardId}，造成 ==${damage}== 点伤害${
          insight ? "（弱点命中 ×2）" : ""
        }${ignoreDef ? "（无视护甲）" : ""}。`,
      );
      if (shardId === "italic") next = { ...next, dodge: true };
      if (shardId === "sup" && rand(1, 100) <= 50) {
        const extra = damageToMonster(state, state.atk, ignoreDef);
        next = { ...next, monsterHp: next.monsterHp - extra };
        log.push(`^上标^ 追击！追加 ${extra} 点伤害。`);
      }
    } else {
      // 普通攻击
      const damage = damageToMonster(state, state.atk + rand(0, 2), false);
      next = { ...next, monsterHp: next.monsterHp - damage };
      log.push(`你挥剑攻击，造成 ${damage} 点伤害。`);
    }

    if (next.monsterHp <= 0) {
      return afterCombat({ ...next, log });
    }

    const turned = monsterTurn(next, log);
    next = turned.state;
    if (next.hp <= 0) {
      return {
        ...next,
        phase: "gameover",
        monster: null,
        log,
        scene: [
          "### 你倒下了",
          "",
          `> 你在第 **${state.floor}** 层被 ${monster.name} 击败。`,
          "",
          "```text",
          ...log.slice(0, 8),
          "```",
          "",
          "> [!TIP] 提示：进入战斗前多收集碎片，`==高亮==` 能让伤害翻倍。",
        ].join("\n"),
        choices: [{ label: "再来一局", action: "restart" }],
      };
    }

    next = { ...next, log };
    return { ...next, scene: combatMarkdown(next), choices: combatChoices(next) };
  }

  // ---- 探索动作 ----
  if (action === "chest:open") {
    const roll = rand(1, 100);
    if (roll <= 45) {
      const remain = SHARDS.filter((s) => !state.shards.includes(s.id));
      if (remain.length > 0) {
        const gained = pick(remain);
        const atkBonus = gained.id === "bold" ? 2 : 0;
        const next: DungeonState = {
          ...state,
          shards: [...state.shards, gained.id],
          atk: state.atk + atkBonus,
          gained: [...state.gained, gained.id],
          log: [`箱中是 ${gained.markdown}（${gained.name}）。`],
          choices: [{ label: "继续深入", action: "descend" }],
        };
        return {
          ...next,
          scene: exploreMarkdown(
            next,
            "宝箱开启",
            `> 你得到语法碎片：${gained.markdown}\n>\n> ${gained.desc}`,
          ),
        };
      }
      const next: DungeonState = {
        ...state,
        ink: state.ink + 2,
        log: ["箱中是两瓶墨水。"],
        choices: [{ label: "继续深入", action: "descend" }],
      };
      return { ...next, scene: exploreMarkdown(next, "宝箱开启", "> 你得到两瓶墨水。") };
    }
    if (roll <= 75) {
      const next: DungeonState = { ...state, hp: Math.max(0, state.hp - 5), log: ["陷阱触发！"] };
      if (next.hp <= 0) {
        return {
          ...next,
          phase: "gameover",
          scene:
            "### 你倒下了\n\n> 箱子上的 `<!-- 危险 -->` 并不是玩笑。\n\n> [!TIP] 提示：宝箱有风险，血量低时不妨先找篝火。",
          choices: [{ label: "再来一局", action: "restart" }],
        };
      }
      return {
        ...next,
        log: [...next.log, "暗箭从箱底射出，你受到 5 点伤害。"],
        choices: [{ label: "继续深入", action: "descend" }],
        scene: exploreMarkdown(next, "陷阱！", "> 你受到 **5** 点伤害。"),
      };
    }
    const next: DungeonState = { ...state, log: ["箱子是空的。"], choices: [{ label: "继续深入", action: "descend" }] };
    return { ...next, scene: exploreMarkdown(next, "空箱子", "> 里面只有一张写着「TODO」的纸条。") };
  }

  if (action === "chest:leave") {
    const next: DungeonState = {
      ...state,
      log: ["你决定不碰它。"],
      choices: [{ label: "继续深入", action: "descend" }],
    };
    return { ...next, scene: exploreMarkdown(next, "谨慎前行", "> 你绕过了箱子，安全第一。") };
  }

  if (action === "rest:sleep") {
    const next: DungeonState = {
      ...state,
      hp: state.maxHp,
      log: ["你睡了一觉，体力恢复。"],
      choices: [{ label: "继续深入", action: "descend" }],
    };
    return { ...next, scene: exploreMarkdown(next, "休息完毕", "> HP 已回满。") };
  }

  if (action === "rest:sharpen") {
    const next: DungeonState = {
      ...state,
      atk: state.atk + 1,
      log: ["你把武器磨得锋利。"],
      choices: [{ label: "继续深入", action: "descend" }],
    };
    return {
      ...next,
      scene: exploreMarkdown(next, "磨剑", "> 攻击力 +1（永久）。"),
    };
  }

  if (action === "book:read") {
    const next: DungeonState = {
      ...state,
      hp: Math.min(state.maxHp, state.hp + 5),
      log: ["你读了一段文字，心情平静。"],
      choices: [{ label: "继续深入", action: "descend" }],
    };
    return {
      ...next,
      scene: exploreMarkdown(
        next,
        "阅读",
        `> 你恢复了 5 点 HP。\n\n> [!NOTE] 书页上写着：${state.lastBookLine}`,
      ),
    };
  }

  if (action === "book:leave") {
    const next: DungeonState = {
      ...state,
      log: ["你合上了书。"],
      choices: [{ label: "继续深入", action: "descend" }],
    };
    return { ...next, scene: exploreMarkdown(next, "离开", "> 有些答案要留到以后再读。") };
  }

  return state;
}

/** 通关结算：在进入第 6 层（即击败 Boss）时调用 */
export function finishRun(state: DungeonState): DungeonState {
  const next: DungeonState = {
    ...state,
    phase: "victory",
    monster: null,
    log: [...state.log, "文档之心安静下来，排版重新变得整齐。"],
    choices: [{ label: "再来一局", action: "restart" }],
  };
  return {
    ...next,
    scene: [
      "# 🏆 通关",
      "",
      "> 你击败了格式混乱之兽，==文档之心== 恢复了平静。",
      "> 所有标题重新对齐，列表整整齐齐，代码块闪着微光。",
      "",
      "**本局收获**",
      "",
      shardListMarkdown(next),
      "",
      "```text",
      ...next.log.slice(0, 10),
      "```",
      "",
      "---",
      "",
      "> [!NOTE] 作者寄语",
      "> 「把复杂留给自己，把简单留给用户。」",
      "> —— **Master Wang（王大师）**",
      "",
      "感谢你把 MasterEdit 玩到这里。它还有很多细节可以打磨，欢迎到软件中心反馈。",
    ].join("\n"),
  };
}

export const BOSS_FLOOR = 5;
export const TOTAL_FLOORS = 5;
