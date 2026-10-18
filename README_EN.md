<div align="center">
  <img src="assets/logo.png" width="76" alt="ShiWenXunBiao icon">
  <h1>ShiWenXunBiao · 识文寻表</h1>
  <p><strong>Select text on screen. Find it across an Excel workbook, instantly.</strong></p>
  <p>Local Windows OCR · Offline · Full-row output · Always on top · Custom shortcut</p>
  <p>
    <a href="README.md">中文</a> · <b>English</b>
  </p>
</div>

<p align="center">
  <img alt="Windows 10/11" src="https://img.shields.io/badge/Windows-10%20%7C%2011-0078D4?logo=windows&logoColor=white">
  <img alt="Python 3.10+" src="https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white">
  <img alt="MIT License" src="https://img.shields.io/badge/License-MIT-yellow.svg">
  <img alt="Local OCR" src="https://img.shields.io/badge/OCR-Local%20%26%20Offline-14B8A6">
</p>

![ShiWenXunBiao UI](assets/screenshot.png)

> Stop copying names or IDs from a web page and pressing `Ctrl + F` inside a huge Excel file. Drag a box, and the row appears.

---

## Why

Excel lookups often start outside Excel: a name on a webpage, an order number in a chat, an ID in another system. Copy → switch to Excel → `Ctrl + F` → paste → search is slow.

**ShiWenXunBiao** turns that into one screen selection:

1. Drag a rectangle around the text you need.
2. Windows OCR reads it locally.
3. The app searches every visible sheet in the workbook.
4. The complete row is shown, and you can jump straight to the cell in Excel.

Everything runs locally. **No screenshots, spreadsheets, or text are uploaded.**

---

## Features

- **Select-to-search** — global hotkey or button, then drag.
- **Local OCR** — Windows 10/11 built-in Chinese/English OCR.
- **Whole workbook** — `.xlsx`, `.xlsm`, and `.csv`, all visible sheets.
- **Full-row output** — every non-empty field in the row, with one-click copy; no truncation.
- **Smart matching** — strips labels such as `Name:` / `Student ID:` and ranks exact, prefix, contains, and fuzzy matches.
- **Jump to Excel** — double-click a result to select the cell in Microsoft Excel.
- **Always on top** — on by default, toggleable, remembered between runs.
- **Customizable hotkey** — record a new combination; conflicts are detected and the previous hotkey is kept.
- **Multi-line OCR** — switch between recognized lines with shortcut chips.
- **High-DPI aware** — consistent screenshot coordinates and layout at 125% / 150%.
- **Private by design** — no network requests, telemetry, or cloud services.

---

## Quick Start

### Option 1: Download a release

1. Download the latest `识文寻表-vX.Y.Z.zip` from **Releases**.
2. Extract it, keeping all files in the same folder.
3. Double-click `识文寻表.exe`.
4. The bundled `示例数据.xlsx` opens automatically.

> `识文寻表.exe` is a small launcher; the actual application is `app.py`. Do not move the exe by itself.
> The release is not a fully standalone Python bundle. See the requirements below.

### Option 2: Run from source

Requirements: Windows 10/11, Python 3.10+ (64-bit with Tkinter).

```powershell
git clone <repository-url>
cd shiwen-xunbiao
python -m pip install -r requirements.txt
python app.py
```

### First run

Select **“张三”** or **“2025010101”** in the sample spreadsheet, then drag over those characters in the running application. The matching row and its full content should appear.

---

## Workflow

1. **Choose a workbook** — click “选择表格”.
2. **Select text** — press `Ctrl + Shift + Q` or click the green button.
3. **Drag a rectangle** around the text, leaving a little margin.
4. The app recognizes the text and searches all sheets automatically.
5. Read the complete row in the “整行内容” panel, or copy it.
6. Double-click a result to select the cell in Excel.

---

## Full-Row Output

- The first result is selected automatically; the detail panel immediately shows **all non-empty fields** in that row.
- Fields are labeled with their column headers.
- “复制整行” copies the complete row without ellipses.
- On short windows the detail panel collapses automatically; click “展开” to open it again.
- The row can also be copied from the result context menu.

---

## Shortcuts and Always-on-Top

| Shortcut | Action |
|---|---|
| `Ctrl + Shift + Q` | Default global screen selection |
| `Ctrl + O` | Choose a workbook |
| `F5` | Reload the current workbook |
| `Enter` | Search from the query field |
| `Ctrl + Enter` | Search |
| `Esc` | Cancel selection / close dialogs |

- Change the global hotkey via “更改” next to the shortcut in the sidebar.
- If a hotkey is occupied, the app reports “冲突” and lets you try another one; the old hotkey is restored if registration fails.
- “置顶：开” in the top-right corner means the window stays above other windows. Click to toggle; the state is remembered.
- The green “框选屏幕文字” button always works, even if the global hotkey is unavailable.

---

## Matching Modes

- **Smart** (default): exact > prefix > contains > fuzzy.
- **Exact**: the cell value must equal the query.
- **Contains**: the cell value must contain the query.

If OCR makes a small mistake, switch to Contains or edit the query manually.

---

## Supported Files

| Format | Notes |
|---|---|
| `.xlsx` | Recommended |
| `.xlsm` | Read-only support; macros are not modified |
| `.csv` | UTF-8 / GBK and other common encodings |
| `.xls`, `.xlsb` | Not supported; save as `.xlsx` first |

- Hidden sheets are skipped.
- Formula cells use Excel’s cached values. Open and save the workbook in Excel if the cache is missing.
- Only the top-left value of a merged cell is indexed.
- Large workbooks take longer to index once; queries are then served from memory.

---

## How It Works

```text
Screen selection
      ↓
Windows.Media.Ocr (local)
      ↓
Text cleanup + query expansion
      ↓
XLSX / CSV parser → in-memory index
      ↓
Scoring: exact / prefix / contains / fuzzy
      ↓
Full-row display / copy / Excel COM jump
```

Stack:

- Python 3.10+ / Tkinter
- Pillow for screenshots and previews
- Windows.Media.Ocr via PowerShell WinRT
- Open XML + zipfile for Excel parsing (no `openpyxl` runtime dependency)
- Win32 `RegisterHotKey` for the global shortcut
- Excel COM for cell navigation

---

## Project Layout

```text
.
├─ app.py                    # UI, state, hotkey settings
├─ excel_index.py            # XLSX / CSV parser and search
├─ ocr_service.py            # Windows OCR wrapper
├─ win_support.py            # hotkey, capture, selection overlay
├─ assets/
│  ├─ ocr_windows.ps1
│  ├─ open_excel.ps1
│  ├─ app.ico
│  ├─ logo.png
│  ├─ screenshot.png
│  └─ social-preview.png
├─ tests/                    # headless core tests
├─ tools/launcher/           # EXE launcher source
├─ docs/                     # GitHub publishing and release notes
├─ .github/                  # CI and GitHub templates
├─ 识文寻表.exe
├─ 示例数据.xlsx
├─ requirements.txt
├─ README.md / README_EN.md
├─ CHANGELOG.md
├─ CONTRIBUTING.md
└─ LICENSE
```

---

## Development

```powershell
python -m pip install -r requirements.txt
python app.py

python -m unittest discover -s tests -v
python -m py_compile app.py excel_index.py ocr_service.py win_support.py
```

Rebuild the launcher with the C# compiler included in .NET Framework:

```powershell
cd tools\launcher
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

---

## Limitations

- Windows 10/11 only.
- Chinese (Simplified) OCR language component must be installed.
- `.xls` / `.xlsb` are not supported.
- Double-click navigation requires desktop Microsoft Excel.
- OCR quality depends on image clarity, font, scaling, and contrast.
- Running from source requires Python 3.10+.
- The exe is a launcher, not a fully self-contained Python distribution.

---

## Roadmap

- [ ] More compact layout for low-resolution screens.
- [ ] Side-by-side full-row comparison.
- [ ] OCR correction for common fields.
- [ ] Pluggable OCR engines.
- [ ] English UI.
- [ ] Automated UI tests.

---

## Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) before opening an issue or pull request.

Please remove names, phone numbers, ID numbers, and other personal data from screenshots and sample workbooks.

If this project helps you, a **Star** is appreciated.

---

## Version History

See [CHANGELOG.md](CHANGELOG.md) for release notes and migration notes.

## License

[MIT License](LICENSE)

Copyright (c) 2026 ShiWenXunBiao contributors

---

<p align="center">
  Made for people who are tired of copy-pasting into Excel search boxes.
</p>
