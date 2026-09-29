# AGENTS.md — MasterEdit（原 mastermd）项目约定

## 项目概况

Windows Markdown 查看与简易编辑器（MasterEdit，原 MasterMD）。技术栈：Tauri 2 + Rust + React 19 + TypeScript + Vite + TailwindCSS 4 + CodeMirror 6。
仓库：https://github.com/bigmasterpang/MasterEdit（2026-09-28 由 MasterMD 改名而来，旧地址会 301 跳转；**代码与 Release 一律推送到新地址**）

### 命名与标识（改名遗留，勿随意"纠正"）

| 标识 | 当前值 | 说明 |
| --- | --- | --- |
| 产品名 / `productName` | `MasterEdit` | 用户可见名称 |
| Cargo 包名 / 可执行文件名 | `MasterEdit` | 包名即二进制名，产物为 `MasterEdit.exe` |
| Rust 库名 | `masteredit_lib` | `src/main.rs` 调用它 |
| `identifier` | `com.masterpang.mastermd` | **故意保留旧值**：它决定应用数据目录（`%APPDATA%\com.masterpang.mastermd`）与安装身份，改掉会让老用户丢设置与 PDF 批注、并产生并存安装 |
| 插件存储文件 | `masteredit-store.json` | 启动时由 `commands/legacy.rs` 从 `mastermd-store.json` 一次性复制迁移，旧文件保留 |
| localStorage 键 | `masteredit.*` | 通过 `src/utils/storage.ts` 的 `readMigratedItem` 兼容读取 `mastermd.*` 旧键 |
| GitHub 仓库 / 徽章链接 | `bigmasterpang/MasterEdit` | 已由 MasterMD 改名，`git remote origin` 已指向新地址 |

## 开发与校验命令

```bash
pnpm install                              # 安装依赖
pnpm typecheck                            # TypeScript 类型检查（提交前必须通过）
pnpm build                                # 仅构建前端
pnpm tauri dev                            # 开发运行
pnpm tauri build                          # 打包（NSIS 安装包）
```

Rust 侧：`cd src-tauri && cargo check --message-format=short`；单元测试：`cargo test --lib`

国内网络需要代理：`$env:HTTP_PROXY="http://127.0.0.1:7890"; $env:HTTPS_PROXY="http://127.0.0.1:7890"`（npm/pnpm 与 ~/.cargo/config.toml 已配置）。

## 提交与发布约定

- 版本号同步修改三处：`package.json`、`src-tauri/Cargo.toml`、`src-tauri/tauri.conf.json`
- 打包产物：`src-tauri/target/release/MasterEdit.exe` 与 `src-tauri/target/release/bundle/nsis/MasterEdit_<版本>_x64-setup.exe`
- 仅在用户明确要求时才执行 git commit / push
- 推送目标：`https://github.com/bigmasterpang/MasterEdit`（凭据已存于 git credential store）
- GitHub Release 通过 REST API 创建，Token 位于 `C:\opencode\github_tokens`（**禁止**写入仓库或输出到日志）
- **发布顺序必须是：commit → push main → 打 tag → push tag → 再创建 Release**。若在推送 tag 前创建 Release，GitHub 会把标签指向当时的默认分支 HEAD，导致 tag 指向错误提交且后续 push 被拒
- 创建 Release 后必须核对 `releases/latest` 的资产名与版本号是否与本地文件一致
- 软件中心只发布了 `installer` 变体（`?variant=portable` 返回 404），应用内更新走安装包静默升级

## 发布到软件中心（单点上传 + 国内节点自动同步）

```powershell
# 打包 + 上传到海外 VM 节点（唯一上传点）
.\tools\publish.ps1 -Mode installer -Build -NotesFile "tools\notes-<版本>.txt"
```

架构（2026-09-28 起）：

- **只上传一次**：发布脚本把安装包提交给海外 VM 主机的 Portal API（`https://master.dapang.wang/api/upload`）。注意 `master.dapang.wang` 与 `vm.dapang.wang` 是**同一台机器**（38.47.108.223）；`vm.dapang.wang` 只提供遗留静态通道（`/updates`、`/apk`、`/masterproxy`），其 `/api` 返回 404，所以上传入口用 master 域名。
- **国内阿里云节点自动同步**：`106.14.225.57` 每 5 分钟由 cron（`/etc/cron.d/portal-sync`）执行 `/opt/app-portal/sync-from-vm.sh`，从 VM 拉取新版本并调用本机 Portal API 重新发布（校验 SHA-256、沿用 10 版保留策略）。脚本规范副本：`C:\opencode\tools\portal-sync-from-vm.sh`。
- 需要跳过等待时，可给 `publish-release.ps1` 传 `-DirectAli`，直接向阿里云节点上传。
- 发布后复核两个节点的版本与 SHA256 是否一致：
  - `https://master.dapang.wang/api/apps/masteredit/windows/latest?variant=installer`
  - `http://106.14.225.57/api/apps/masteredit/windows/latest?variant=installer`

## 办公表格模块（0.18.0 起）

- **CSV / TSV**：走现有文本管线（内容存在 `doc.content`，可编辑、可撤销、按编码保存）；表格视图在 `src/components/Sheet/`（`SheetView` + `SheetGrid`）与 `src/utils/delimited.ts`。视图三态复用全局 `viewMode`：`preview` = 表格、`source` = 源码、`split` = 表格 + 源码。
- **CSV 用惰性行索引**：`createDelimitedTable(text)` 单遍扫描只记录行首偏移（十几 MB / 十几万行约 30ms，偏移表约 1MB），行内容由 `table.rowAt(i)` 按需解析并缓存最近若干行 —— **没有行数上限**，也不为每个单元格建对象。`parseDelimited(text, {maxRows})` 保留给一次性取前 N 行的场景，两条路径结果必须一致（有等价性对拍测试）。
- **xlsx / xls / xlsb / ods 读取**：Rust 侧 `src-tauri/src/commands/office.rs`（calamine）按行窗口解析（`spreadsheet_info` / `spreadsheet_rows`），前端稀疏缓存 + 虚拟滚动；文档模型为 `docType: "spreadsheet"`。写回能力见下一条：0.19.0 起只有 `.xlsx` 可编辑，其余格式与超大文件 `readOnly: true`；**自动保存对表格一律跳过**（整簿重写必须用户显式 Ctrl+S），外部改动只刷新基线并触发重新解析。
- **企业透明加密（亿赛通等）**：工作簿在 `read_workbook_bytes` 中先 `esafenet::decrypt_esafenet` **内存解密**再交给 calamine（`open_workbook_auto_from_rs` + `Cursor`），**绝不写明文临时文件**；`SpreadsheetInfo.encrypted` 供界面显示「已解密」。直接读原始字节会得到 `Could not find EOCD` 这类误导性错误。
- **解析缓存**：`SheetCache` 是按「最近使用」排序的 LRU，受 `CACHE_ENTRIES` 与 `CACHE_CELL_BUDGET`（总单元格预算）双重约束；**单张超大工作表会被保留**，否则百万行工作表每个滚动窗口都要重解析整表。
- **打不开时的提示**：`describe_parse_failure` 按文件头区分「需要密码的 Office 文档 / 旧版 CFB 格式」「HTML 假 Excel」「损坏或被加密软件处理过」，给出可执行建议，而不是抛底层 zip 错误。
- **xlsx 轻量编辑（0.19.0 起）**：写回在 `src-tauri/src/commands/office_write.rs`（umya-spreadsheet 3.1）。要点：**只有 `.xlsx` 可写回**（`SpreadsheetInfo.editable`；xlsm 含宏、xls/xlsb/ods 一律只读）；写回是**整簿重写**：「临时文件 → 用 calamine 校验可读 → `fs::rename` 原子替换」（中途失败原文件不受影响）；`.bak` 备份**默认关闭**（可在设置里开启）；企业加密文件全程内存解密/加密，**不落明文临时文件**；公式不做重算（前端首次保存会确认提示）。前端待提交编辑存在 `DocState.sheetEdits`（唯一数据源，驱动脏标记、关闭确认、会话恢复），撤销历史见下一条，保存走 `fileActions.saveSpreadsheetDoc`。
- **约定**：**不要**把 Office 格式加进 `tauri.conf.json` 的 `fileAssociations`（避免抢占默认打开程序、放大版式保真预期差）；CSV/TSV 走文本管线可自由编辑；xlsx 做「改值 + 行列结构 + 工作表增删改复制」级轻量编辑，**不做样式编辑（字体/颜色/对齐/格式刷）、图表/透视表/条件格式**；公式计算是**只读求值引擎**（`office_formula.rs`，25 个函数 + 整列/整行引用，结果只用于界面显示，文件里仍只写公式）。
- **表格查找 / 统计 / 结构操作（0.20.0 起）**：`src-tauri/src/commands/office_ops.rs` 提供 `spreadsheet_find`（全表扫描，500 条命中封顶）、`spreadsheet_stats`（矩形区域的求和/平均/计数/最值，选区超出一屏时前端改用它）、`spreadsheet_structure`（插入/删除行列、新建/重命名/删除/复制工作表）。结构操作**只改内存里的影子工作簿**（`office_shadow.rs`：快照栈 30 步 / 512MB），界面立刻生效并标记「结构已修改（未保存）」，**按 `Ctrl+S` 才整簿写回**（复用 `office_write::save_pipeline`）；配套命令 `spreadsheet_state` / `spreadsheet_undo` / `spreadsheet_redo` / `spreadsheet_discard`，前端不弹确认框；**插入/删除位置超出已用范围按空操作成功返回**（不报错、不扩张尺寸），仍会报错的是 `count > 100000`、表不存在、重名或非法表名、删最后一张表、非 `.xlsx`；复制工作表是「新建表 + 逐格 `Cell::clone` + 列宽/行高/合并区域」，上限 20 万格。公式**会算**（见下条），但公式引用**不会随结构操作平移**（与 Excel 行为不同，属已知取舍）。读取与解析（含企业加密内存解密、失败提示话术）直接复用 `office.rs` 的 `read_workbook_bytes` / `parse_workbook` / `describe_parse_failure`（`pub(crate)`），不另存一份实现。
- **表格撤销历史（0.20.0 起）**：`src/utils/sheetHistory.ts` 按**文档 id** 保存在模块级 Map 中（不再用组件 ref），切标签、切视图、组件重建后仍可撤销；一个撤销步可含多格（清空选区、粘贴只占一步）。xlsx 的撤销步必须记录**文件里的真实原值**（`SheetEdit` 而非 null），否则保存后撤销会退化成「删掉待提交编辑」而界面毫无变化；文档关闭时用 `clearSheetHistory` 释放。
- **表格交互（0.20.0 起）**：撤销栈 `src/utils/sheetHistory.ts` 覆盖**内容、结构、列宽行高**（布局步骤 key 为 `col:N` / `row:N`，随结构操作平移；一次拖动只记一步，见 `src/hooks/useSheetLayout.ts`）；冻结窗格与手动宽高由父组件经 `freezeRows`/`freezeCols`/`columnWidths`/`rowHeights` 受控传入；查找替换面板在 `src/components/Sheet/SheetFind.tsx`（`Aa`/`W`/`.*`/循环、整表或选区、替换只作用于当前工作表、xlsx 不支持正则），文本匹配与替换的公共实现在 `src/utils/sheetReplace.ts`；工具栏经 `src/utils/sheetCommands.ts` 的命令通道驱动当前激活的表格视图（自动调整、插入公式、冻结）。
- 相关测试：`cd src-tauri && cargo test --lib office`（Excel 序列日期、类型映射、窗口收敛、非 A1 区域坐标、公式提取）。

## 完成后通知

- **不要在项目脚本里发微信通知**：微信通知统一由 DeepSeek Harness 的通知插件负责。共享脚本 `C:\opencode\tools\publish-release.ps1` 内的旧微信卡片已**默认关闭**（需要时传 `-Notify` 才发），`tools/publish.ps1` 另外固定传 `-SkipNotify`；`C:\opencode\tools\send-wechat.ps1` 仅保留作手动调用。

### 发布脚本编码陷阱（务必遵守）

PowerShell 5.1 会把 **无 BOM 的 UTF-8 脚本按 GBK 解析**：脚本里的中文注释可能吞掉行尾换行，导致下一行代码被并入注释而静默失效（曾导致 Release 资产名错误）。因此：

- 发布用 `.ps1` 脚本**只用 ASCII 注释**；中文说明一律写入单独的 UTF-8 文件后用 `-Encoding UTF8` 读取
- 必须含中文的脚本要用「UTF-8 with BOM」保存
- 上传 Release 资产前先确认本地文件名与 `?name=` 参数一致，传完用 API 复核 `releases/latest` 的资产名

## 代码约定

- TypeScript strict；函数组件 + Hooks；全局状态用 Zustand；不使用大型 UI 库
- 关键逻辑写中文注释；仅在必要时新增依赖，优先选择体积更小、行为更稳定的方案
- 前端产物通过 Vite 代码分割按需加载（KaTeX / Mermaid / highlight.js 语言包）
