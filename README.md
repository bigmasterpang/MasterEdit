<div align="center">

# MasterEdit <small>（原 MasterMD）</small>

**为 Windows 打造的轻快全能文档与代码阅读编辑器**

Markdown 读写导出 · PDF 深度批注 · 代码智能导览 · 双栏独立对照 · 色弱无障碍友好

[![Release](https://img.shields.io/github/v/release/bigmasterpang/MasterEdit?style=flat-square&label=%E6%9C%80%E6%96%B0%E7%89%88%E6%9C%AC)](https://github.com/bigmasterpang/MasterEdit/releases)
[![Downloads](https://img.shields.io/github/downloads/bigmasterpang/MasterEdit/total?style=flat-square&label=%E4%B8%8B%E8%BD%BD%E9%87%8F)](https://master.dapang.wang)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11%20x64-0078D4?style=flat-square)](https://master.dapang.wang)
[![Size](https://img.shields.io/badge/%E5%AE%89%E8%A3%85%E5%8C%85-5.1%20MB-2ea44f?style=flat-square)](https://master.dapang.wang)
[![License](https://img.shields.io/badge/license-Freeware-lightgrey?style=flat-square)](#授权与致谢)

![MasterEdit 分屏预览](docs/screenshots/main.png)

</div>

---

## 为什么是 MasterEdit

> 从专注 Markdown 的 **MasterMD** 进化为全能轻快的 **MasterEdit**：
> 告别动辄几百 MB 的重型 IDE 与臃肿阅读器，用不到 **5 MB** 的体积，把 **「Markdown 写作导出、PDF 阅读批注、代码函数级导览、双栏并排对照」** 四件事做到极致顺手。

| 维度 | MasterEdit | 常见 Electron 编辑器 / 重型阅读器 |
| --- | --- | --- |
| 安装包体积 | **约 5.1 MB** | 80 – 300 MB |
| 冷启动速度 | **约 0.3 秒** | 2 – 6 秒 |
| 空闲内存占用 | **约 30 MB** | 200 – 600 MB |
| 多格式一体 | **Markdown + PDF 批注 + 代码导览** | 需分别安装多个软件 |
| 双栏对照 | **每栏独立缩放、独立导航历史、支持同文档对照** | 全局缩放联动或不支持同文件分栏 |
| 更新方式 | **应用内一键校验并自更新（双节点容灾）** | 手动下载重装 |

---

## 核心特性

### 🧭 代码智能阅读与导览（Code Viewer & Navigation）
- **多语言符号引擎**：内置 Rust、TypeScript / JavaScript、Python、Go、C / C++、Java、C# 等主流语言的函数、类、结构体、接口与方法解析
- **跳转到定义（`F12` / `Ctrl+左键`）**：在任意函数或类型引用处一键定位至原定义位置，支持跨文件工作区检索
- **速览原函数实现（`Alt+F12`）**：无需离开当前上下文，内联弹窗**精准提取并仅展示目标函数体**，配备独立滚动条，支持一键在左/右分栏并排打开
- **查找所有引用（`Shift+F12`）**：汇总展示符号在当前文件及工作区内的所有定义与调用位置，点击即可预览或跳转
- **顶部符号面包屑导航栏**：实时显示 `文件路径 › 所属类/模块 › 当前函数`，点击可下拉快速切换同文件任意函数
- **分栏独立导航历史栈**：左/右分栏各自独立维护「返回上一位置（`Alt+←`）/ 前进下一位置（`Alt+→`）」，跨栏打开函数绝不污染原分栏历史或引发焦点乱跳

### 📑 PDF 深度阅读与无损批注
- **划词高亮与批注合一**：选中文字即可添加多色高亮（自动合并同行分片消除重叠色块）、支持右键随时切换高亮颜色，或直接附加文字批注
- **自由便签卡片**：页面任意位置插入便签，智能避让防遮挡，支持随时拖拽移动与编辑
- **全层级书签大纲 + 批注侧栏**：完整解析媲美 Adobe Acrobat 的多级嵌套书签树；侧栏集中聚合所有高亮、批注与便签，点击直达对应页码与位置
- **无损持久化与污染修复**：批注独立保存并支持写回 PDF 标准注释层；内置历史受污染 PDF 物理高亮深度剥离与还原能力
- **全操作撤销与重做**：高亮、换色、批注、便签、页面顺/逆时针旋转均支持 `Ctrl+Z` / `Ctrl+Y` 多步撤回

### 🪟 真·双栏独立对照工作区
- **左右双栏自由组合**：支持左看 PDF 文献、右写 Markdown 笔记，或左看函数调用处、右看原函数实现（支持同一文档在左右两栏同时打开不同位置）
- **智能分栏右键菜单**：根据标签当前所在栏位动态切换「在右侧分栏打开 / 在左侧分栏打开」
- **窗口分栏独立缩放**：按住 `Ctrl + 鼠标滚轮` 仅放大/缩小当前鼠标所在分栏的文档（编辑区、预览区、PDF 均按分栏独立记忆缩放比例），绝不干扰另一栏

### ♿ 无障碍体验（红绿色弱友好模式）
- 在 **设置（`Ctrl+,`）→ 外观与无障碍** 中一键开启 **「色弱友好模式（红绿色弱优化）」**：
  - **Okabe-Ito 国际通用无障碍代码高亮**：彻底避开易混淆的红-绿/橙-绿对比，采用**洋红紫关键字（配半粗体双重线索）**、**天青蓝字符串**、**暖琥珀金数字/常量**、**波浪下划线错误提示**
  - **PDF 高亮「色彩 + 底部线型」双重编码**：将三色高亮升级为高对比度的 **明黄（实线底边）/ 天青蓝（虚线底边）/ 洋红紫（双线底边）**，工具栏、右键菜单与侧栏同步显示无障碍色名与线型标识
  - **带中文色名的主题强调色**：设置面板强调色选择器采用「色点 + 中文色名胶囊按钮」，无需仅凭裸眼辨色

### 📊 办公表格查看（CSV / TSV / Excel）
- **CSV / TSV 表格视图**：自动识别分隔符（逗号 / 分号 / 制表符），兼容引号字段、字段内换行与 BOM；`Ctrl+E` 在 **表格 / 源码 / 表格+源码** 之间切换，需要改数据时切到源码即可编辑保存
- **Excel 只读查看**：内置 Rust 解析引擎（calamine），支持 `.xlsx / .xls / .xlsb / .ods`；多工作表标签、列标与行号和 Excel 完全对齐、单元格内容 / 公式栏、百万行级虚拟滚动不卡顿
- **大文件也流畅**：十几 MB、十几万行的 CSV 采用「只扫行偏移、按需解析行」的惰性索引，没有行数上限，滚动内存占用远低于一次性解析
- **支持企业加密文档**：亿赛通等透明加密的 Excel / CSV 会在内存中解密后查看，不产生明文临时文件；打不开时会明确告知原因（需要密码、旧版格式、假 Excel 等）
- **安全只读**：表格文档绝不写回原文件，避免格式被破坏；需要编辑时右键「用系统默认程序打开」交给 Excel / WPS
- **不劫持文件关联**：安装时不会抢占 `.xlsx / .docx` 的默认打开程序，随时可以用回 Office / WPS

### 📖 Markdown 阅读与 ✍️ 顺手编辑
- **GFM 与扩展语法全支持**：表格、任务列表、提示块 `> [!NOTE]`、高亮 `==文字==`、下划线 `++文字++`、上下标、YAML Front Matter 元数据卡片
- **公式与图表按需加载**：KaTeX 数学公式、Mermaid 流程图/时序图/状态图按需延迟加载
- **三视图无缝切换**：预览 / 源码 / 分屏（`Ctrl+E` 一键循环），分屏双向滚动精准同步
- **多标签页 + 会话自动恢复**：意外关闭不丢草稿，启动自动还原上次打开的 Markdown、代码、PDF 与表格文档及阅读位置

### 📤 四格式一键导出
| 格式 | 特点 |
| --- | --- |
| **HTML** | 单文件自包含，内联样式与本地图片，随手分享 |
| **PNG** | 2 倍超清长图渲染，公式、图表、代码高亮完整保留 |
| **Word (.docx)** | 标题/列表/表格/代码块/图片完整转换，纯前端生成标准 OOXML |
| **PDF** | 直连 WebView2 原生矢量打印引擎，无需额外打印驱动 |

---

## 下载与使用

**方式一：软件中心下载（推荐，含国内高速节点）**

👉 <https://master.dapang.wang>

**方式二：GitHub Releases**

| 文件 | 说明 |
| --- | --- |
| `MasterEdit_x.y.z_x64-setup.exe` | **安装版（推荐）**：注册开始菜单、桌面快捷方式与 `.md` / `.pdf` 等文件关联 |
| `MasterEdit_x.y.z_x64.exe` | **便携版**：双击即用，无需安装，可放任意目录 / U 盘 |

> 需要 WebView2 运行时：Windows 11 自带；Windows 10 会在安装时自动引导安装。

---

## 快捷键速查

| 操作 | 快捷键 | | 操作 | 快捷键 |
| --- | --- | --- | --- | --- |
| 打开文件 / 文件夹 | `Ctrl+O` | | 跳转到函数定义 | `F12` / `Ctrl+左键` |
| 保存 / 另存为 | `Ctrl+S` / `Ctrl+Shift+S` | | 速览原函数实现 | `Alt+F12` |
| 切换视图（预览/源码/分屏） | `Ctrl+E` | | 查找所有引用 | `Shift+F12` |
| 返回上一位置 / 前进下一位置 | `Alt+←` / `Alt+→` | | 当前分栏独立缩放 | `Ctrl+滚轮` |
| 查找 / 替换 | `Ctrl+F` / `Ctrl+H` | | 加粗 / 斜体 | `Ctrl+B` / `Ctrl+I` |
| 撤销 / 重做（含 PDF 批注） | `Ctrl+Z` / `Ctrl+Y` | | 标题 H1–H6 | `Ctrl+1` … `Ctrl+6` |
| 关闭标签 / 切换标签 | `Ctrl+W` / `Ctrl+Tab` | | 设置 / 快捷键面板 | `Ctrl+,` / `F1` |

---

## 技术栈

<div align="center">

**Tauri 2** · **Rust** · **React 19** · **TypeScript** · **Vite** · **TailwindCSS 4** · **CodeMirror 6** · **PDF.js** · **calamine** · **markdown-it** · **highlight.js** · **DOMPurify**

</div>

- **前后端分离**：文件读写、目录扫描、文件监听、PDF 原生清理、安全下载与打印走 Rust；界面、编辑器与多格式渲染走 WebView2
- **极致体积**：KaTeX / Mermaid / PDF 引擎 / 代码语言包全部通过 Vite 代码分割按需加载
- **安全优先**：`html: false` + DOMPurify 双重净化，外部链接交由系统默认浏览器安全打开

---

## 开发与构建

```bash
# 环境要求：Node ≥ 20、pnpm ≥ 9、Rust stable ≥ 1.77、VS Build Tools 2022
pnpm install
pnpm tauri dev        # 开发运行
pnpm typecheck        # TypeScript 类型检查（提交前必过）
pnpm build            # 仅构建前端
pnpm tauri build      # 打包 EXE 与 NSIS 安装包
```

国内网络可先设置代理：

```powershell
$env:HTTP_PROXY="http://127.0.0.1:7890"; $env:HTTPS_PROXY="http://127.0.0.1:7890"
```

发布到软件中心（双节点自动分发 + 微信通知）：

```powershell
.\tools\publish.ps1 -Mode installer -Build -NotesFile "<utf8更新说明文件>"
```

---

## 项目结构

```text
src/
├── components/    工具栏 / 编辑器(含代码面包屑与速览弹窗) / PDF批注视图 / 预览 / 大纲与侧栏 / 搜索 / 设置
├── hooks/         渲染管线、滚动同步、文件监听、主题与色弱无障碍同步、自动保存、快捷键
├── stores/        Zustand：文档与分栏状态、设置持久化、搜索、弹窗、在线更新
└── utils/         代码符号分析与导航历史栈、Markdown 渲染、导出管线、持久化

src-tauri/src/
├── lib.rs         插件注册、单实例唤醒、窗口关闭拦截、启动参数解析
└── commands/      文件读写、目录树、最近文件、图片、文件监听、PDF 打印、在线更新
```

---

## 授权与致谢

- 作者：**Master Wang（王大师）** · Master 系列软件
- 问题反馈：<https://github.com/bigmasterpang/MasterEdit/issues>
- 软件中心：<https://master.dapang.wang>

<div align="center">

**如果 MasterEdit 让你少装了几个几百 MB 的软件，欢迎点个 ⭐ Star**

</div>
