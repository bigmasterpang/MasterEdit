# MasterEdit · Word（.docx）查看功能 —— 进度与问题交接文档

> 面向接手修复的 AI/工程师。本文**自包含**：读完即可定位代码、复现问题、跑通验收。
> 最后更新：0.21.4 发布后（提交 `049016a`）。

---

## 一、项目概况

| 项 | 值 |
| --- | --- |
| 产品 | MasterEdit（原 MasterMD）—— Windows 上的 Markdown / 文本 / PDF / Excel / **Word** 查看与轻量编辑工具 |
| 技术栈 | Tauri 2 + Rust + React 19 + TypeScript(strict) + Vite + TailwindCSS 4 + CodeMirror 6 + Zustand |
| 仓库 | `https://github.com/bigmasterpang/MasterEdit`（工作目录 `C:\opencode\mastermd`） |
| 当前版本 | **0.21.4**（`package.json` / `src-tauri/Cargo.toml` / `src-tauri/tauri.conf.json` 三处同步） |
| 发布渠道 | GitHub Release + 海外 VM 门户 `master.dapang.wang` + 阿里云节点 `106.14.225.57`（自动同步） |

**定位（重要）**：Word 文档**只做查看、不做编辑**。需要编辑时用工具栏「打开方式」交给系统里的 Word / WPS。渲染目标是**结构忠实**，不是像素级还原。

---

## 二、Word 查看功能当前能做什么（0.21.0 → 0.21.4）

| 能力 | 状态 |
| --- | --- |
| 打开 `.docx`（含**亿赛通等企业透明加密**，内存解密、不落明文临时文件） | ✅ |
| 段落排版：字体/字号/颜色/粗斜体/下划线/删除线/上下标/对齐/左右缩进/首行悬挂缩进/段前段后/行距（倍数与固定值） | ✅ |
| 样式继承链：`docDefaults → 段落/字符样式(basedOn 逐级) → 段落标记 rPr → 直接格式`；含**主题字体**（`theme1.xml`） | ✅ |
| 多级编号：中文数字/罗马数字/带圈数字/项目符号，`%N` 展开、按 `numId` 独立计数 | ✅ |
| 表格：`gridSpan` 横合并、`vMerge` 纵合并（前端配对算 rowspan）、列宽、逐边边框、底纹、表头行、单元格内多段落与嵌套表格 | ✅ |
| 图片：懒加载 + `(path, media)` 缓存 + 失败/超时/重试 | ✅ |
| **文本框**（封面/表单常见）：内部段落、表格、图片递归渲染，含填充与边框 | ✅ |
| **页眉页脚**：逐页渲染（多节时从最后一节往前找可用的 default 引用） | ✅ |
| **实时页码**：`PAGE` / `NUMPAGES` 域按实际分页显示（不是缓存值） | ✅ |
| **分页视图**：按文档真实 `w:sectPr`（纸张尺寸 + 页边距 + 横竖版）分页，白纸卡片 + 页间空隙 + 纸外页码；可切「连续」 | ✅ |
| 形状：VML（`v:line`/`v:rect`/`v:roundrect`/`v:oval`/`v:hr`）与 DrawingML（`prst=line|rect|roundRect|ellipse`）画成真实线框 | ✅ |
| 段落边框 `w:pBdr`（合同里"空段落 + 下边框"画的横线） | ✅ |
| 查找（`Ctrl+F`，块级命中 + 滚动定位 + 高亮）、大纲侧栏（按 `outlineLevel`）、缩放（`Ctrl+滚轮`） | ✅ |
| 复制：正文可选中；「复制全部文本」（带列表编号）；表格右键「复制为制表符文本」 | ✅ |
| 占位兜底：SmartArt / 图表 / OLE / 组合图形 / 自由曲线 / EMF·WMF / 空文本框 → **带中文说明的占位卡片，绝不静默丢失** | ✅ |
| `.docx` 进入「常用打开」与资源管理栏白名单 | ✅（0.21.3） |

**语料库实测（29 份真实文档，见第四节）**：全部解析成功、**占位块 0**、零 panic。

---

## 三、代码地图与数据契约

### 3.1 后端（Rust）

| 文件 | 说明 |
| --- | --- |
| `src-tauri/src/commands/office_docx.rs` | **全部 DOCX 逻辑**（约 6400 行，含测试）。解密 → 解包 → 解析块模型 → 序列化 |
| `src-tauri/src/lib.rs` | `invoke_handler` 注册 5 个命令；`pub mod commands;`（集成测试需要） |
| `src-tauri/tests/docx_corpus.rs` | 语料库完备性测试（284 行，见第四节） |
| `src-tauri/src/commands/office.rs` | 复用其 `read_raw_and_plain`（透明解密入口） |

**命令（前端契约，已冻结）**：

```
document_info(path)                          → DocumentInfo
document_xml(path)                           → String（原始 word/document.xml）
document_blocks(path, from, count)           → BlockPage    ← 主接口，窗口化
document_find(path, query, matchCase)        → FindHit[]     ← 注意 JS 侧参数名是 camelCase
document_media(path, media)                  → String（data URL；仅允许 word/media/ 下，防目录穿越）
```

`document_blocks` / `document_find` / `document_media` 都是 `#[tauri::command(async)]`，**完全无状态**（`from` 原样回显，供前端丢弃过期窗口）。解析结果按「路径 + 大小 + mtime」缓存最近 2 个文档。

### 3.2 块模型（前端 TypeScript 联合类型）

```ts
type DocBlock =
  | ({ kind: "paragraph" } & DocParagraph)
  | ({ kind: "table" } & DocTable)
  | ({ kind: "image" } & DocImage)
  | { kind: "pageBreak" }
  | ({ kind: "shape" } & DocShape)
  | ({ kind: "textBox" } & DocTextBox)
  | { kind: "unsupported"; label: string; detail: string };
```

关键字段（完整定义见 `src/types/index.ts`）：

- `DocParagraph`：`runs` / `text` / `align` / `indentLeftPt` / `indentFirstLinePt`（负数=悬挂）/ `spaceBeforePt` / `spaceAfterPt` / `lineSpacing{kind,value}` / `list{numId,level,ordered,prefix,format,suffix}` / `outlineLevel`（**大纲只能用这个**）/ `pageBreakBefore` / `sectionBreak` / `borders{top,bottom,left,right}`
- `DocRun`：`text` / `bold` / `italic` / `underline` / `strike` / `font`（西文）/ `fontEastAsia`（中文，**font-family 优先用它**）/ `sizePt` / `color` / `highlight` / `vertAlign` / **`field: "PAGE" | "NUMPAGES" | null`**
- `DocTable` / `DocTableRow` / `DocTableCell`：`gridSpan` / `vMerge("none"|"restart"|"continue")` / `shading` / `vAlign` / `borders` / `blocks`（单元格内块）
- `DocImage`：`media` / `name` / `alt` / `widthPx` / `heightPx`
- `DocShape`：`shape("line"|"rect"|"roundRect"|"ellipse")` / `xPt,yPt,widthPt,heightPt` / `lineWidthPt` / `lineColor` / `fillColor` / `dash` / `vertical`
- `DocTextBox`：`blocks`（内部块）/ `xPt,yPt,widthPt,heightPt` / `fillColor` / `borderColor` / `borderWidthPt` / `wrap`
- `BlockPage`：`total` / `from` / `blocks` / `encrypted` / **`page: PageGeometry | null`**（纸张几何）/ **`header` / `footer`: `DocBlock[] | null`**

### 3.3 前端（`src/components/Docx/`，共 12 个文件约 3900 行）

| 文件 | 职责 |
| --- | --- |
| `DocxView.tsx` (517) | 容器：工具条、图片缓存与请求去重、右键菜单、`Ctrl+滚轮` 缩放、视图模式 |
| `DocxBlocks.tsx` (587) | 虚拟滚动：高度索引（前缀和）、可见窗口、测量修正、滚动锚点补偿、分页画布 |
| `useDocxBlocks.ts` (207) | 数据层：窗口取块 + 后台预取 + generation 丢弃过期响应 |
| `DocxBlock.tsx` (699) | 单块渲染：段落（run 级格式 + 列表前缀）、图片、形状、文本框、分页线、占位卡片 |
| `DocxTable.tsx` (219) | `<table>`：colgroup / colSpan / rowSpan / 边框 / 底纹 / 表头 / 嵌套 |
| `docxGrid.ts` (293) | 纯函数：rowspan 配对、表格→TSV、块→纯文本、大纲收集、占位块合并计划 |
| `docxStyle.ts` (650) | 纯函数：**字段→CSS 映射 + 高度估算**（估算与渲染必须共用这里的函数） |
| `docxPages.ts` (221) | 纯函数：分页装页（块→页映射）、页几何解析 |
| `docxRender.ts` (62) | 渲染上下文（含滚动容器 ref、缩放、图片缓存 key） |
| `docxCopy.ts` (48) / `DocxFind.tsx` (270) / `DocxOutline.tsx` (71) | 剪贴板 / 查找 / 大纲 |

---

## 四、测试与验收基础设施（接手后先跑一遍）

### 4.1 命令

```powershell
# 前端
pnpm typecheck          # 必须 0 错误
pnpm build              # 必须先构建：部分断言会读 dist/assets/*.css

# 后端
cd src-tauri
cargo test --lib                     # 当前 183 passed
cargo test --test docx_corpus        # 当前 3 passed（29 份真实文档 + 2 个合成样本）
cargo check --all-targets            # 必须 0 警告
```

### 4.2 语料库（**最重要的一条**）

`src-tauri/tests/docx_corpus.rs` 默认扫描两个目录（可用环境变量 `MASTEREDIT_DOCX_CORPUS` 覆盖，多目录用 `;` 分隔）：

```
Z:\D\mywork\05_会议汇报\每周汇报\2026                                    (25 份加密周报)
Z:\D\mywork\08_项目总结\AI加持下的家庭能源管理系统                        (4 份：企业标准 / 查新合同 / 实验室对接方案 / 电弧与负荷辨识)
```

它会：① 断言每份文档要么解析成功、要么给出**可操作的中文错误**（绝不 panic）；② 校验 `document.xml` 良构；③ 统计并打印**每份文档的顶层块/形状/占位块**，并断言"占位块必须有 label"；④ 两个**合成样本**覆盖真实语料没有的路径（只写 `firstLineChars` 的字符单位缩进、只写 `beforeAutospacing` 的自动段间距）。

**目录不存在时自动跳过**（CI/别的机器不会失败）。**新增真实文档请直接放进这两个目录** —— 这是本项目发现问题的唯一有效手段。

### 4.3 前端 harness（`%TEMP%\sheetcheck\`，仓库外）

```powershell
cd $env:TEMP\sheetcheck
node docx.mjs        # 当前 391 passed —— Word 视图的全部断言（jsdom + 真实组件 + mock invoke）
node test.mjs        # 123  表格
node edit.mjs        # 840  表格编辑
node view.mjs        # 407  表格视图
node view-task3.mjs  # 375  表格任务 3
```

### 4.4 真实内核探针（**验滚动/懒加载/溢出只能靠它**）

```powershell
node probe-docx.mjs <dump.json> <截图.png> ok     # 把真实 DocxView 打进 bundle，在无头 Edge 里跑真滚动、真 IO
node docx-shot.mjs                                # 渲染 + 逐块量高度（找出"估算 vs 实测"差异）
node docx-real.mjs                                # 头部注释里有"用真实块模型 dump 出 JSON"的配方
```

探针会输出 `overlaps`（块重叠数）、`pageOverflows`（内容压出版心数）、`horizontalOverflow`、`backToTop`、图片加载状态等指标。**本项目的排版问题几乎都是靠这些数字定位的**（见第六节）。

---

## 五、已知问题清单（按优先级，供接手修复）

### P0 —— 用户已反馈、尚未修复或需要确认

| # | 问题 | 复现 | 线索 / 建议 |
| --- | --- | --- | --- |
| 1 | **文本框的纵向节奏与 Word 不一致**：封面标题在 Word 里居中、下方留大片空白；我们按文档流紧凑排列 | `AI加持下的家庭能源管理系统企业标准.docx` 封面 | 该文档 4 个文本框全是 `x=0, y=0, wrap=none`（内联），**不能简单绝对定位**（会全部叠在页顶）。若要还原，需要按文本框在文档流中的顺序 + 各自声明高度做间距分配，或读取 `wp:anchor`（若存在）。见 `docxStyle.ts` 的 `resolveTextBoxBox` |
| 2 | **`N20 + 34 个空格 + Q/NDB` 那行的换行位置**与 Word 不完全一致 | 同上文档第 4 个文本框 | 已做"连续 ≥3 空格按西文字体渲染"（`splitRunSegments`），但仍非逐字符取字体。Word 是逐字符按 Unicode 范围选 `w:ascii`/`w:eastAsia` |
| 3 | **页眉页脚不随节切换**：整篇只用一套 | 多节文档（企业标准有 4 节） | 目前"从最后一节往前找第一个有 default 的引用"。完整实现需要按节切分页面并各用各的页眉页脚；`w:titlePg`（首页不同）与奇偶页也未区分 |
| 4 | **「复制全部文本」里的页码是缓存值**（例如 `2`），不是实时页码 | 任意带 PAGE 域的文档 | 这是**有意**的（复制=文档文本语义，与渲染分页解耦）。若要改，在 `docxGrid.ts` 的 `blocksToPlainText` 里需要页序信息 |
| 5 | **浮动图片按内联显示**，不做文字环绕 | 带 `wp:anchor` 环绕的文档 | 后端已给 `wrap`（文本框有，图片暂无），前端未做环绕排版 |
| 6 | **EMF / WMF 矢量图**能取到字节但浏览器渲染不了 → 占位卡片 | Word 里粘贴 Excel 图表 | 需转码（后端或前端 WASM），成本较高 |

### P1 —— 已知保真度缺口（有意不做，可评估）

| # | 缺口 | 说明 |
| --- | --- | --- |
| 7 | `w:contextualSpacing`（同样式段落之间不加空） | 是渲染规则、需新增契约字段；段落间距目前按 XML 实际数值渲染 |
| 8 | `.doc`（Word 97-2003 二进制） | 不支持，给出"请用 Word/WPS 另存为 .docx"提示 |
| 9 | `.docm`（含宏） | 按只读处理，不写回 |
| 10 | 修订痕迹 | 按**最终态**显示（`w:ins` 展开、`w:del` 跳过），不提供接受/拒绝 |
| 11 | 嵌套文本框的缩进处理 | 文本框里的文本框按同样规则渲染，缩进未做特殊处理 |
| 12 | 形状锚在文字流中间 | 形状是独立块，不会精确嵌进那一段的文字流（对"空段落 + 下划线/方框"的合同形态没问题） |
| 13 | 表格单元格内的图片按单元格内宽算显示盒 | 与估算同口径，但可能小于 Word 的显示尺寸 |

### P2 —— 需要用户继续实测确认（0.21.4 刚修，尚无用户反馈）

- 横向溢出（0.21.4 修）：探针实测三份文档 `horizontalOverflow: 0`
- `王诚0911.docx` 翻页回顶（0.21.4 修）：探针 40 步细粒度上滚零次被顶回
- `王诚0918.docx` 图片加载（0.21.4 修）：8 秒超时 + 失败原因 + 重试按钮
- 企业标准封面：日期行是否仍为一行、标题字号（后端 22pt 加粗居中，已核对一致）

---

## 六、踩过的坑（**务必遵守，都是真实事故**）

### 6.1 渲染层：三条不变式

1. **估算高度必须等于渲染高度** —— 所有新块类型都必须让 `docxStyle.ts` 的估算函数与实际 DOM 高度一致，否则会出现"下一块压上来"（用户已报过 3 次同类问题）。真浏览器探针是唯一的验证手段。
2. **字符宽度表必须覆盖所有宽字符** —— 事故：中文全角引号 `“ ”`（U+201C/D）不在宽字符表里，按半角 0.52em 计，一句话少算 30.7px → 估成 1 行、实际 2 行 → 下一块压上来 31px。**任何新增的标点/符号都要检查**。
3. **字体影响宽度** —— 事故：39 个连续空格按中文字体 0.5em 计（西文 Times New Roman 只有 0.25em），多出 196px。连续空白必须按西文字体渲染（`splitRunSegments`）。
4. **段落容器的字号/字体要跟随段落主字号/主字体** —— 否则容器 strut 会撑大行盒（事故：22pt 标题差 2~5px）。

### 6.2 React：两个真实 bug 的教训

5. **`useEffect` 依赖里不能放每次渲染都新建的对象** —— 事故：图片加载 effect 依赖 `ctx` → 反复重跑 → cleanup 置 `cancelled = true` → 请求回来时 setState 被丢弃 → **永久卡在"加载中"**。改用 ref 取最新值。
6. **可选字段判断不能用 `!== null`** —— 事故：`run.field !== null` 在**老数据**（字段不存在 → `undefined`）下把**所有 run** 判成域 run → 全部 `inline-block` → 两端对齐把正文拉出空隙并多占一行 → 6 处重叠。必须写 `=== "PAGE" || === "NUMPAGES"`。

### 6.3 滚动与溢出

7. **纸张永不产生横向滚动** —— `overflow: clip`（不是 `hidden`：`clip` 不创建滚动容器，不会吃掉纵向滚动）。文本框"布局宽按版心、内容层可溢出到纸张边界"。
8. **测量修正后要补偿滚动锚点**，但补偿值必须夹取，否则用户"翻到第二页回不去第一页"（已发生）。

### 6.4 工程环境（本项目特有）

9. **PowerShell 5.1 会把无 BOM 的 UTF-8 脚本按 GBK 解析** —— 发布脚本只用 ASCII 注释，中文说明写单独 UTF-8 文件。**绝不要用 PowerShell 的 `Get-Content`/`Set-Content` 改含中文的 `.mjs`/`.ts` 源文件**（会按 GBK 解码毁掉中文）—— 用编辑器工具。
10. **PowerShell 的 `$PWD` 与 .NET 的进程工作目录可能不一致** —— `[System.IO.File]::ReadAllText("相对路径")` 会失败，**一律用绝对路径**。
11. **发布顺序必须是**：commit → push main → 打 tag → push tag → 创建 Release（否则 tag 指向错误提交）。发布后核对三渠道的版本与 SHA256。
12. **不要给 `tauri.conf.json` 加 Office 文件关联**（避免抢占默认打开程序）。

---

## 七、接手建议（如果要继续提升保真度）

按"性价比"排序：

1. **先让用户用 0.21.4 实测**（第五节 P2 的四项），拿到具体差异再动手 —— 本项目所有有效修复都来自真实文档实测，自造样本几乎覆盖不到。
2. **文本框纵向位置**（P0-1）：这是封面/合同类文档观感的最大剩余差异。
3. **逐字符字体解析**（P0-2）：让"ASCII 走西文字体、CJK 走中文字体"的宽度与换行彻底对齐 Word。
4. **按节切换页眉页脚 + 首页不同**（P0-3）。
5. **浮动图片环绕**（P0-5）：需要后端补 `wrap` 与锚点坐标。

**每次改动后必须跑**：`pnpm typecheck` + `pnpm build` + `cargo test --lib` + `cargo test --test docx_corpus` + `node docx.mjs`（+ 表格四个 harness），并用真实内核探针确认 `overlaps = 0`、`pageOverflows = 0`、`horizontalOverflow = 0`。
