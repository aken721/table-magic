<div align="center">

# 🧰 TableMagic Add-in | 万象插件

**The Excel Plugin That Does the Heavy Lifting**

[🇨🇳 中文](./README.md) · [![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT) · [![Version](https://img.shields.io/badge/Version-v2.5.2-blue.svg)](https://gitee.com/aken721/table-magic/releases) · [![Platform](https://img.shields.io/badge/Platform-Office%20%2B%20WPS-orange.svg)]() · [![Stars](https://img.shields.io/badge/Stars-⭐%20Give%20a%20Star-yellow.svg)](https://gitee.com/aken721/table-magic)

[📥 Download](https://gitee.com/aken721/table-magic/releases) · [🐛 Issues](https://gitee.com/aken721/table-magic/issues) · [📧 Custom Dev](mailto:aken721@163.com)

</div>

---

### Introduction

**TableMagic Add-in** is a completely **free and open-source** Excel add-in built with VSTO technology, compatible with both Microsoft Office and WPS.

Unlike ordinary toolboxes that only provide basic spreadsheet operations, this add-in focuses on **real-world, high-frequency workplace pain points** — wrapping complex tasks like invoice processing, bulk email sending, contract generation, scheduled automation, and AI-powered natural language operations into **a few simple clicks**, empowering non-technical office workers to benefit from automation.

### ✨ Key Features

#### 📊 Spreadsheet Enhancement

**Data Split & Merge**
- **Split Sheets/Workbooks**: Split data into separate sheets or workbooks by condition with one click; header rows are preserved
- **Merge Sheets/Workbooks**: Quickly consolidate data from multiple sheets or workbooks vertically
- **Batch Import from Folder**: Merge specified sheets from all workbooks in a folder into the current workbook

**Table Structure Transformation**
- **2D to 1D Table (Unpivot)**: Convert pivot/cross-tabulated tables back to flat detail records — date columns auto-formatted, numeric columns stripped of thousands separators; essential pre-processing for data analysis
- **Payroll Table → Pay Slips**: Split a consolidated payroll sheet into individual pay slips per employee with configurable spacing rows

**Data Extraction & Cleaning**
- **Regex Data Extraction**: 9 built-in presets (numbers / English text / Chinese characters / URLs / ID numbers / email / phone / IP address / custom regex) — extract exactly what you need from any cell
- **Auto-Renumber Sequence Column**: After split/merge operations, automatically detect and renumber the sequence column continuously

**Worksheet Management**
- **Excel Table of Contents** (bidirectional):
  - Auto-generate a hyperlinked index page from all existing worksheets — one-click navigation to any sheet
  - Batch-create blank worksheets from a name list — rapidly scaffold multi-sheet report structures
- **Batch Create Blank Sheets**: Create 1–15 sheets at once with a custom name prefix
- **Batch Delete Sheets**: Delete multiple worksheets by rule
- **Spotlight**: Highlights the entire row and column of the active cell; supports 6 color themes — no more losing your place
- **Export to PDF**: Current sheet, all sheets to single PDF, or all sheets to separate PDFs with custom page size and scaling

#### 📸 Screenshot (New in v2.5.2)

- **Auto Window Detection**: Move the mouse to automatically highlight the window under the cursor with a blue border and transparent preview — similar to Windows/WeChat screenshot
- **Drag-to-Select**: Hold left mouse button and drag to select any rectangular area
- **Double-Click to Capture Window**: Double-click to capture the currently highlighted window
- **Two Modes**: Current window screenshot (Excel stays visible) / Hidden window screenshot (minimize Excel first)
- Screenshot is automatically copied to clipboard and inserted into the active worksheet at the selected cell, scaled proportionally to fit 20 columns × 10 rows
- Multi-monitor support; Esc/right-click to cancel, Enter to confirm

#### 🧾 Invoice Management (Unique Feature)
- **XML Electronic Invoices**: Batch-import key invoice data into Excel
- **PDF/Image Invoice OCR**: Powered by PaddleOCR — batch-extract data from scanned invoices or photos
- Say goodbye to manual invoice entry!

#### 📝 Batch Word Document Generation (Unique Feature)
- Set placeholders in a Word template; each row in Excel automatically generates a separate document
- **Supports image insertion** — write the image file path in Excel, and it's automatically embedded in the Word document
- Auto-detects `.doc` and `.docx` format
- Typical use cases: bulk contracts, offer letters, employee certificates, material labels

#### 📧 Bulk Email Sending (Unique Feature)
- Send **personalized content and attachments** to different recipients based on an Excel table
- Built-in **HTML email editor** (custom-built rich text control) — compose and send beautifully formatted HTML emails
- No more manual copy-paste!

#### ⏰ Scheduled Tasks (Unique Feature)
- Visual task scheduling in Windows Task Scheduler style
- Supports: **CMD commands / VBA macro code / BAT scripts / Python scripts / EXE programs**
- Trigger types: one-time, daily, weekly, monthly, custom interval
- Configure and save scheduled tasks without ever leaving Excel

#### 📁 Batch File Management
- **Bulk Rename**: Load directory files into Excel, edit new names freely in the table, then apply in bulk
- **Bulk Move/Delete**: Supports subdirectories
- Paths and filenames support hyperlinks — click to open directly

#### 📈 Chart Enhancement
- **Interactive HTML Charts**: Generate web-based interactive charts (powered by ECharts) with zoom, filter, and dynamic display
- **GIF Animated Charts**: Create animated chart presentations — drop them into PowerPoint for instant impact
- **Pareto Chart**: Auto-calculate cumulative percentages and generate a combined bar+line Pareto analysis chart with one click
- **Word Cloud**: Count cell word frequencies and generate word clouds; export as SVG vector or PNG bitmap

#### 🔲 QR Code & Barcode
- Generate **colorful QR codes** (custom foreground/background color, logo embedding supported)
- **Multi-column merge generation**: Concatenate multiple columns in `key:value;` format and batch-generate QR codes
- Generate **Code128 barcodes**
- Recognize QR codes and barcodes from images or webcam, and extract data into Excel

#### 🎵 Background Music
- Play background music right inside Excel — no more silence during tedious data entry
- Supports mp3 / wav / flac / aiff / wma / aac / mp4 and other mainstream formats

#### 🗄️ Direct Database Export (Unique Feature)
- Supports 6 major databases: **MySQL / SQL Server / Access / SQLite / PostgreSQL / Oracle**
- Keyword fuzzy search for table names; real-time DataGridView preview
- Long numeric values auto-converted to text format — prevents scientific notation truncation
- Access / SQLite support file selection via double-click browse
- Export database tables to Excel in one click — no SQL, no coding required

#### 🌐 API Data Write-In (Unique Feature)
- Enter an API endpoint URL and data is automatically written into the current Excel sheet — no Python or VBA required
- 3 authentication modes: **No Auth / Username+Password (Basic Auth) / API Key** (custom header name and value prefix supported)
- URL placeholder substitution for dynamic, parameterized endpoint addresses
- Auto-adapts to TLS 1.0 / 1.1 / 1.2 for compatibility with legacy servers
- Typical use cases: internal enterprise systems, weather/finance/government open data feeds

#### 🤖 AI Chat & Natural Language Excel Operations (Major Update in v2.5)
- Supports 14 mainstream LLM providers: DeepSeek, Qwen, Kimi, ChatGLM, ChatGPT, Claude, and more
- Supports local models via **Ollama / LM Studio / vLLM / llama.cpp** (data stays on your machine)
- Automatically detects model capabilities, intelligently switches between Function Calling and Prompt Engineering modes
- **Skills plug-in architecture**: Load third-party extensions via the standard ISkill interface — 24 built-in skills covering workbooks, worksheets, cells, ranges, formatting, charts, pivot tables, financial analysis, and more
- Just describe what you want in plain language — AI handles the Excel operation

#### 🔌 Universal MCP Skill (New in v2.5.1)

The add-in's 111 Excel operation tools are packaged as a **universal MCP Skill** that can be installed and used in standalone Agents such as openclaw, hermes, workbuddy, qclaw, and Trae Work — no Excel installation required.

- **Protocol**: MCP (Model Context Protocol) 2024-11-05, stdio transport
- **Engine**: ClosedXML (no Excel installation needed)
- **Features**: 111 tools for workbook/worksheet/cell/range/format/chart/analysis/database/API/email/file/QR/invoice/regex/TOC/schedule/PDF operations
- **Performance**: Single-process persistent MCP mode — 10 sheet creations in just 4ms; compound tools complete multi-step operations in a single call

**Installation**:

1. **Download pre-built package** (recommended): Download `table-magic-skill.zip`, extract it to the Agent's skill installation directory, and the Agent will automatically discover and register this skill via `skill.md` and `mcp-config.json`; or drag `table-magic-skill.zip` into the Agent's chat window and instruct the Agent to install the skill
2. **Build from source**: `git clone` this repository, run the `pack-skill.ps1` script to build and generate `table-magic-skill.zip`, then extract the package to the Agent's skill installation directory

See [TableMagic.Cli/README.md](./TableMagic.Cli/README.md) for details.

### 📥 Installation

1. Go to the [GitHub](https://github.com/aken721/table-magic/releases) or [GitCode](https://gitcode.com/akin721/table-magic/releases) Releases page and download the `.msi` installer
2. After installation, open Excel — the **"Wanxiang Plugin"** tab will appear automatically in the ribbon (versions prior to v2.5.1.0 use the tab name "Intelligent Toolbox")
3. The installer includes WPS registry support — works directly in WPS Office

### 🔧 System Requirements

| Item | Requirement |
|:---|:---|
| OS | Windows 7 or later |
| Office | Microsoft Office 2013+ or WPS |
| .NET | .NET Framework 4.7.2+ |
| Invoice OCR | Requires separate OCR model package |

### 💡 Who Is This For?

- 📊 **Finance/Accounting**: Batch invoice entry, financial report automation
- 👥 **HR/Admin**: Batch contract generation, bulk notifications, file renaming
- 📈 **Operations/Sales**: Data consolidation, report automation
- 🏭 **Procurement/Supply Chain**: Database import, bulk price processing
- Anyone spending too much time on repetitive tasks in Excel

### 🤝 Custom Development

Every business has unique needs. If the built-in features don't fully cover your workflow, feel free to reach out for **custom development**:

📧 **aken721@163.com**

### 📜 License

MIT License — Completely free, commercial use allowed, just keep the copyright notice.

Copyright (c) 2023 Ken He / Hainan Nuowo Digital Technology Co., Ltd.

---

<div align="center">

**If this project helps you, please give it a ⭐ Star!**

[Gitee](https://gitee.com/aken721/table-magic) · [GitHub](https://github.com/aken721/table-magic) · [GitCode](https://gitcode.com/akin721/table-magic)

📋 [View Full Changelog →](./CHANGELOG.md)

</div>
