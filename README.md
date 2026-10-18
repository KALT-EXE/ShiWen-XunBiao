<div align="center">
  <img src="assets/logo.png" width="76" alt="识文寻表图标">
  <h1>识文寻表 · ShiWenXunBiao</h1>
  <p><strong>框选屏幕上的文字，自动在 Excel 全表中查找并定位。</strong></p>
  <p>Windows 本地 OCR · 离线运行 · 整行输出 · 窗口置顶 · 自定义快捷键</p>
  <p>
    <b>中文</b> · <a href="README_EN.md">English</a>
  </p>
</div>

<p align="center">
  <img alt="Windows 10/11" src="https://img.shields.io/badge/Windows-10%20%7C%2011-0078D4?logo=windows&logoColor=white">
  <img alt="Python 3.10+" src="https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white">
  <img alt="MIT License" src="https://img.shields.io/badge/License-MIT-yellow.svg">
  <img alt="Local OCR" src="https://img.shields.io/badge/OCR-Local%20%26%20Offline-14B8A6">
</p>

![识文寻表界面](assets/screenshot.png)

> 不再对着屏幕上的姓名或编号，在几万行 Excel 里按 `Ctrl + F`。框一下，结果直接出现。

---

## 目录

- [为什么做这个](#为什么做这个)
- [核心特性](#核心特性)
- [快速开始](#快速开始)
- [使用流程](#使用流程)
- [整行内容输出](#整行内容输出)
- [快捷键与窗口置顶](#快捷键与窗口置顶)
- [搜索模式](#搜索模式)
- [支持的文件](#支持的文件)
- [工作原理](#工作原理)
- [项目结构](#项目结构)
- [开发与测试](#开发与测试)
- [已知限制](#已知限制)
- [路线图](#路线图)
- [贡献](#贡献)
- [License](#license)

---

## 为什么做这个

很多 Excel 查询场景并不发生在 Excel 里：姓名在网页上、编号在聊天窗口里、单号在另一个系统里。传统做法是复制、切回 Excel、按 `Ctrl + F`、粘贴、再找。

**识文寻表**把这条路径缩短成一次框选：

- 在任意界面框住文字；
- 调用 Windows 自带的本地 OCR 识别；
- 自动在整本 Excel 中查找；
- 直接展示完整整行，并可跳回 Excel 定位。

OCR、检索、索引全部在本机完成，**不联网、不上传截图和表格**。

---

## 核心特性

| 特性 | 说明 |
|---|---|
| **框选即查** | 全局快捷键或按钮框选屏幕区域，自动识别并查找 |
| **本地 OCR** | 使用 Windows 10/11 自带中英文 OCR，无需云服务 |
| **整本工作簿** | 一次检索全部可见工作表，支持 `.xlsx`、`.xlsm`、`.csv` |
| **整行输出** | 选中结果后，下方直接显示整行全部非空字段，一键复制，不做省略 |
| **智能匹配** | 自动去掉「姓名：」「学号：」等标签，按完全 / 开头 / 包含 / 近似匹配排序 |
| **Excel 定位** | 双击结果，在 Microsoft Excel 中打开并选中对应单元格 |
| **窗口置顶** | 默认始终显示在其他窗口上方，可随时切换并记住状态 |
| **自定义快捷键** | 支持现场录制新组合键，冲突时给出提示并保留原快捷键 |
| **多行识别** | 一次框选多行时，用「快捷」标签切换要检索的内容 |
| **高 DPI** | 适配 Windows 125% / 150% 等缩放，截图坐标与界面一致 |
| **零上传** | 无网络请求，无遥测，无云端依赖 |

---

## 快速开始

### 方式一：下载直接运行

1. 在 GitHub 的 **Releases** 页面下载最新的 `识文寻表-vX.Y.Z.zip`。
2. 解压到任意文件夹，**保持文件完整**。
3. 双击 `识文寻表.exe`。
4. 首次启动会自动打开 `示例数据.xlsx`。

> `识文寻表.exe` 是轻量启动器，实际逻辑在 `app.py`。请连同同目录文件一起移动，不要只移动 exe。

### 方式二：从源码运行

环境要求：Windows 10/11、Python 3.10+（64 位，带 Tkinter）。

```powershell
git clone <repository-url>
cd shiwen-xunbiao
python -m pip install -r requirements.txt
python app.py
```

也可以直接双击 `备用启动.cmd`。

### 第一次体验

框选示例 Excel 里的 **“张三”** 或 **“2025010101”**，即可看到查询结果和完整整行内容。

---

## 使用流程

1. **选择表格**：点击左侧「选择表格」，载入 `.xlsx` / `.xlsm` / `.csv`。
2. **框选文字**：按默认快捷键 `Ctrl + Shift + Q`，或点击绿色「框选屏幕文字」。
3. **拖拽选区**：在屏幕上拖出矩形，给文字四周留一点边距，松开鼠标。
4. **自动识别查找**：识别完成后自动检索全部工作表。
5. **查看整行**：结果下方显示完整整行内容，可复制。
6. **定位 Excel**：双击某条结果，在 Excel 中选中对应单元格。

---

## 整行内容输出

- 结果表默认自动选中第一条，下方「整行内容」面板立即显示该行**所有非空字段**。
- 字段按表头名称排列，例如：

  ```text
  学号：2025010101 | 姓名：张三 | 班级：园艺2026级1班 | 专业：园艺 |
  宿舍：紫荆1-302 | 手机号：13800138001 | 报到日期：2026-09-01 | 状态：在校 | 备注：班长
  ```

- 点击「复制整行」可复制完整内容，不带省略号。
- 小窗口下会自动收起详情面板以保证结果列表可用，点击「展开」即可恢复。
- 右键结果也可以复制整行或单元格地址。

---

## 快捷键与窗口置顶

| 快捷键 | 功能 |
|---|---|
| `Ctrl + Shift + Q` | 默认全局框选识别 |
| `Ctrl + O` | 选择 Excel 表格 |
| `F5` | 重新载入当前表格 |
| `Enter` | 在查找框中开始查找 |
| `Ctrl + Enter` | 开始查找 |
| `Esc` | 取消框选 / 关闭说明或设置窗口 |

- 默认快捷键可以在左侧「更改」中重新录制；如果与系统或其他软件冲突，程序会显示「冲突」并允许更换。
- 新快捷键注册失败时，会保留原快捷键，不会让你失去可用入口。
- 右上角「置顶：开」表示窗口始终在其他窗口上方；点击可切换，状态自动记住。
- 即使全局快捷键不可用，绿色「框选屏幕文字」按钮始终可以点击。

---

## 搜索模式

- **智能匹配**（默认）：完全匹配 > 开头匹配 > 包含匹配 > 近似匹配，适合大多数场景。
- **精确匹配**：单元格内容必须与检索词一致。
- **包含匹配**：单元格中包含检索词即可。

如果 OCR 有少量错字，可切换「包含匹配」或手动修改查找内容。

---

## 支持的文件

| 格式 | 说明 |
|---|---|
| `.xlsx` | 推荐 |
| `.xlsm` | 支持读取，不会修改宏 |
| `.csv` | 自动识别 UTF-8 / GBK 等常见编码 |
| `.xls` / `.xlsb` | 暂不支持，请先另存为 `.xlsx` |

- 隐藏工作表不会被检索。
- 公式单元格优先使用 Excel 保存的缓存值；如果刚修改公式，请先用 Excel 打开并保存一次。
- 合并单元格只索引左上角内容。
- 文件越大，首次建立索引越慢；之后查询在内存中完成。

---

## 工作原理

```text
┌──────────────┐
│ 屏幕框选截图  │
└──────┬───────┘
       ▼
┌─────────────────────┐
│ Windows.Media.Ocr   │  本地中英文 OCR
└──────┬──────────────┘
       ▼
┌─────────────────────┐
│ 文本清洗 / 查询词扩展 │  “姓名：张三” → “张三”
└──────┬──────────────┘
       ▼
┌─────────────────────┐
│ XLSX / CSV 解析索引  │  共享文本 · 样式 · 行数据
└──────┬──────────────┘
       ▼
┌─────────────────────┐
│ 打分与排序           │  完全 / 开头 / 包含 / 近似
└──────┬──────────────┘
       ▼
┌──────────────────────────────┐
│ 整行展示 / 复制 / Excel 定位  │
└──────────────────────────────┘
```

技术栈：

- **Python 3.10+ / Tkinter**：桌面界面
- **Pillow**：截图与图像预览
- **Windows.Media.Ocr**：系统 OCR（通过 PowerShell WinRT 调用）
- **Open XML / zipfile**：无需 `openpyxl` 即可读取 Excel
- **Win32 `RegisterHotKey`**：全局快捷键
- **Excel COM**：双击定位单元格

---

## 项目结构

```text
.
├─ app.py                    # 主界面、状态管理、快捷键设置
├─ excel_index.py            # XLSX / CSV 解析、索引、搜索
├─ ocr_service.py            # Windows OCR 封装与图像预处理
├─ win_support.py            # 全局快捷键、截图、框选浮层
├─ assets/
│  ├─ ocr_windows.ps1        # WinRT OCR 调用
│  ├─ open_excel.ps1         # Excel COM 定位
│  ├─ app.ico
│  ├─ logo.png
│  ├─ screenshot.png
│  └─ social-preview.png     # GitHub 社交预览图
├─ tests/                    # 核心逻辑单元测试
├─ tools/launcher/           # EXE 启动器源码与构建脚本
├─ docs/                     # GitHub 发布与 Release 文案
├─ .github/                  # CI、Issue/PR 模板
├─ 识文寻表.exe              # 轻量启动器
├─ 示例数据.xlsx
├─ 创建桌面快捷方式.vbs
├─ 备用启动.cmd
├─ requirements.txt
├─ README.md
├─ README_EN.md
├─ CHANGELOG.md
├─ CONTRIBUTING.md
└─ LICENSE
```

---

## 开发与测试

```powershell
python -m pip install -r requirements.txt

# 运行
python app.py

# 核心逻辑测试（不需要 GUI）
python -m unittest discover -s tests -v

# 语法检查
python -m py_compile app.py excel_index.py ocr_service.py win_support.py
```

启动器源码位于 `tools/launcher/`，可使用 Windows 自带的 .NET Framework C# 编译器重新生成：

```powershell
cd tools\launcher
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

GitHub 发布、About 文案、Topics 和 Release 步骤见 [docs/GITHUB_UPLOAD.md](docs/GITHUB_UPLOAD.md)。

---

## 已知限制

- 仅支持 **Windows 10/11**；依赖 Windows OCR 和 Win32 API。
- 系统需要安装**中文（简体）OCR**语言组件。
- 不支持旧版 `.xls` 和 `.xlsb`，请先另存为 `.xlsx`。
- 双击定位依赖桌面版 Microsoft Excel；未安装 Excel 时仍可查询和复制。
- OCR 准确率受截图清晰度、字体、缩放和背景对比度影响；框选时请留边距。
- 运行源码需要本机安装 Python 3.10+；如果换电脑，请按「快速开始」配置。
- 发布包中的 exe 只是启动器，不是包含 Python 运行时的独立安装包。

---

## 路线图

- [ ] 更紧凑的低分辨率结果布局。
- [ ] 多结果整行对比视图。
- [ ] 常见字段的 OCR 自动纠错（学号 / 姓名 / 班级）。
- [ ] 可选的其他 OCR 引擎适配。
- [ ] 英文界面。
- [ ] 更完整的自动化 UI 测试。

欢迎在 Issues 中补充使用场景和优先级。

---

## 贡献

欢迎提交 Bug、功能建议或 Pull Request。开始之前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)。

- Bug 反馈请附 Windows 版本、Python 版本和脱敏后的复现信息。
- 不要上传真实姓名、手机号、身份证号等敏感数据。
- 如果你觉得这个项目有帮助，欢迎点一个 **Star**。

---

## 版本历史

版本变化与迁移说明见 [CHANGELOG.md](CHANGELOG.md)。

## License

本项目基于 [MIT License](LICENSE) 开源。

Copyright (c) 2026 ShiWenXunBiao contributors

> 如果你要用于商业项目，请保留许可证声明；如需其他授权方式，可以自行调整 `LICENSE`。
