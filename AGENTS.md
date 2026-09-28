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
