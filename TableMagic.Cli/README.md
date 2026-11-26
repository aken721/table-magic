# TableMagic CLI & MCP Server

TableMagic 通用型 Skill，可作为 MCP 服务器在 openclaw、hermes、workbuddy、qclaw、trae work 等 Agent 中安装调用。无需安装 Excel，基于 ClosedXML 实现。

## 功能概览

| 类别 | 能力 | 工具数 |
|------|------|--------|
| 基础操作 | 工作簿/工作表信息获取、批量调用、复合操作 | 7 |
| 工作簿/工作表 | 创建、打开、保存、复制、重命名、删除、冻结窗格 | 16 |
| 单元格/区域 | 读写值、公式、批量赋值、复制、清除 | 10 |
| 格式排版 | 字体、颜色、边框、合并、换行、行高列宽、插入删除行列 | 11 |
| 图表/透视 | 创建图表、数据透视表 | 2 |
| 数据分析 | 统计分析、财务比率、利润率 | 4 |
| 数据处理 | 分表、并表、批量导删、转置、工资条 | 9 |
| 数据库 | SQL Server/MySQL/PostgreSQL/SQLite 连接、查询、导出 | 4 |
| REST API | 调用接口获取数据写入 Excel | 2 |
| 增强图表 | 词云、动态图、对比图、帕累托图、直方图、箱线图 | 6 |
| 邮件群发 | SMTP 配置、单发、群发、预览 | 5 |
| 文件操作 | 批量重命名、复制、移动、删除、文件夹管理 | 8 |
| 二维码 | 生成 QR 码/条形码、扫描识别（ZXing.Net + SkiaSharp） | 4 |
| 发票识别 | XML 发票导入、OCR 图片识别（PaddleOCR）、批量导入 | 5 |
| 正则提取 | 正则提取、替换、验证 | 3 |
| 目录页 | 生成目录表、根据目录建表、更新超链接 | 3 |
| 定时任务 | 创建/删除/启用/禁用/立即执行定时任务 | 6 |
| Word 文档 | 批量生成 Word 文档、预览 | 2 |
| PDF 导出 | 单表/全簿/批量/区域导出为 PDF（QuestPDF） | 4 |

**共 111 个工具** | **版本 2.5.2.1**

## 安装

### 方式一：ZIP 安装包（推荐，无需 .NET SDK）

1. 下载 `table-magic-skill.zip`，解压到 Agent 的技能安装目录，Agent 会自动通过 `skill.md` 和 `mcp-config.json` 发现并注册本技能；或者将 `table-magic-skill.zip` 拖入 Agent 对话框，发出安装本技能的指令后，Agent 自动安装并注册
2. 无需任何手工配置，Agent 自动识别技能并启动 MCP 服务

### 方式二：从源码构建安装包

```bash
git clone <repo-url>
cd TableMagic
./pack-skill.ps1
```

构建完成后生成 `table-magic-skill.zip`，按方式一安装即可。

### 方式三：dotnet 全局工具（需要 .NET 8 SDK）

```bash
dotnet pack TableMagic.Cli
dotnet tool install --global --add-source ./nupkg TableMagic.Cli
tablemagic mcp
```

### 方式四：从源码直接运行

```bash
dotnet build TableMagic.Cli
dotnet run --project TableMagic.Cli -- mcp
```

## 使用方式

### 1. MCP 服务器模式（供 Agent 调用，推荐）

Agent 安装技能后自动以 MCP 模式启动，无需手工操作。

```bash
tablemagic mcp
```

### 2. 命令行直接调用（仅用于调试，性能较差）

```bash
# 列出所有工具
tablemagic list-tools

# 调用工具
tablemagic call create_workbook --args '{"fileName":"test.xlsx"}'
tablemagic call set_cell_value --args '{"fileName":"test.xlsx","row":1,"column":1,"value":"Hello"}'

# 批量执行
tablemagic batch calls.json
```

## Agent 自动发现

本技能包含 `skill.md` 和 `mcp-config.json`，支持 Agent 自动发现和注册，无需手工配置 MCP。各 Agent 的技能安装目录：

| Agent | 技能安装目录 |
|-------|-------------|
| OpenClaw | `~/.openclaw/skills/table-magic/` |
| Hermes | `~/.hermes/skills/table-magic/` |
| WorkBuddy | `~/.workbuddy/skills/table-magic/` |
| QClaw | 项目根目录 `skills/table-magic/` |
| Trae Work | 项目根目录 `.trae/skills/table-magic/` |

将 `table-magic-skill.zip` 解压到对应目录，或将 zip 拖入 Agent 对话框即可完成安装。

## 参数引导机制

当用户 prompt 不明确或缺少必需参数时，工具会返回结构化的缺失参数信息，Agent **必须**向用户追问缺失参数，不要猜测或使用默认值：

1. **必需参数缺失**：返回 `[MISSING_PARAM]` 标记，列出缺失参数名称和描述
2. **参数值模糊**：如用户说"创建一个表格"但未指定文件名，应追问"请提供工作簿文件名"
3. **一次只追问最关键的缺失参数**，避免一次列出过多问题

## 性能优化

| 调用方式 | 10次建表耗时 | 说明 |
|----------|------------|------|
| CLI `call` 模式 | ~8,500 ms | 每次启动新进程，极慢，**禁止使用** |
| MCP 模式（进程内） | ~4 ms | 单进程持久运行，**推荐** |
| 复合工具 `create_workbook_with_sheets` | ~61 ms | 一次调用完成多步操作，**最优** |
| `batch_call` | ~395 ms | 一次MCP调用执行多个工具 |

**最佳实践**：
- 优先使用复合工具（`create_workbook_with_sheets`、`write_table_data`、`read_table_data`）
- 批量操作使用 `batch_call`
- 禁止使用 `tablemagic.exe call` 命令行逐个调用

## 工具详细说明

### 基础操作 (ExcelBase)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `get_worksheet_names` | 获取所有工作表名称 | 无 | fileName |
| `get_open_workbooks` | 获取打开的工作簿列表 | 无 | 无 |
| `get_workbook_metadata` | 获取工作簿元数据 | 无 | fileName |
| `batch_call` | 批量执行多个工具调用（一次MCP调用执行多步操作） | calls | 无 |
| `create_workbook_with_sheets` | 创建工作簿并批量创建多个工作表 | fileName | sheetNames, sheetNamePrefix, sheetCount |
| `write_table_data` | 一次性写入完整表格（表头+数据） | fileName, sheetName, headers, data | startRow, startColumn, headerStyle, autoFit |
| `read_table_data` | 一次性读取完整表格 | fileName, sheetName | range, includeHeaders |

### 工作簿操作 (ExcelWorkbook)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `create_workbook` | 创建新工作簿 | fileName | sheetName |
| `open_workbook` | 打开工作簿 | fileName | 无 |
| `close_workbook` | 关闭工作簿 | 无 | fileName |
| `save_workbook` | 保存工作簿 | 无 | fileName |
| `save_workbook_as` | 另存为新文件 | newFileName | fileName |
| `delete_workbook` | 删除工作簿文件 | fileName | 无 |

### 工作表操作 (ExcelSheet)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `activate_worksheet` | 激活/切换工作表 | sheetName | fileName |
| `create_worksheet` | 创建新工作表 | sheetName | fileName |
| `rename_worksheet` | 重命名工作表 | oldSheetName, newSheetName | fileName |
| `delete_worksheet` | 删除工作表 | sheetName | fileName |
| `copy_worksheet` | 复制工作表 | sourceSheetName, targetSheetName | fileName |
| `move_worksheet` | 移动工作表位置 | sheetName, position | fileName |
| `set_worksheet_visible` | 设置工作表可见性 | sheetName, visible | fileName |
| `get_worksheet_index` | 获取工作表索引 | sheetName | fileName |
| `freeze_panes` | 冻结窗格 | 无 | fileName, sheetName, row, column |
| `unfreeze_panes` | 取消冻结窗格 | 无 | fileName, sheetName |

### 单元格操作 (ExcelCell)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `set_cell_value` | 设置单元格值 | row, column, value | fileName, sheetName |
| `get_cell_value` | 获取单元格值 | row, column | fileName, sheetName |
| `set_cell_formula` | 设置单元格公式 | cellAddress, formula | fileName, sheetName |
| `get_cell_formula` | 获取单元格公式 | cellAddress | fileName, sheetName |

### 区域操作 (ExcelRange)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `set_range_values` | 批量设置区域值 | rangeAddress, data | fileName, sheetName |
| `get_range_values` | 获取区域值 | rangeAddress | fileName, sheetName |
| `copy_range` | 复制区域 | sourceRange, targetRange | fileName, sheetName |
| `clear_range` | 清除区域内容 | rangeAddress | fileName, sheetName, clearType |
| `get_used_range` | 获取已使用范围 | 无 | fileName, sheetName |
| `get_last_row` | 获取最后行号 | 无 | fileName, sheetName |

### 格式设置 (ExcelFormat)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `set_cell_format` | 设置单元格格式 | rangeAddress | fileName, sheetName, fontColor, backgroundColor, fontSize, bold, italic, horizontalAlignment, verticalAlignment |
| `set_border` | 设置边框 | rangeAddress, borderType | fileName, sheetName, lineStyle |
| `merge_cells` | 合并单元格 | rangeAddress | fileName, sheetName |
| `unmerge_cells` | 取消合并 | rangeAddress | fileName, sheetName |
| `set_cell_text_wrap` | 设置自动换行 | rangeAddress, wrap | fileName, sheetName |
| `set_row_height` | 设置行高 | rowNumber, height | fileName, sheetName |
| `set_column_width` | 设置列宽 | columnNumber, width | fileName, sheetName |
| `insert_rows` | 插入行 | rowIndex | fileName, sheetName, count |
| `insert_columns` | 插入列 | columnIndex | fileName, sheetName, count |
| `delete_rows` | 删除行 | rowIndex | fileName, sheetName, count |
| `delete_columns` | 删除列 | columnIndex | fileName, sheetName, count |

### 图表与分析 (ExcelChart/ExcelPivot/ExcelAnalysis/ExcelFinance)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `create_chart` | 创建图表 | dataRange | fileName, sheetName, chartType, title |
| `create_pivot_table` | 创建数据透视表 | sourceRange, pivotSheetName | fileName, sheetName, rowFields, columnFields, valueFields |
| `analyze_data` | 分析数据 | range | fileName, sheetName |
| `get_range_statistics` | 获取统计信息 | range | fileName, sheetName |
| `calculate_financial_ratio` | 计算财务比率 | revenueRange, costRange | fileName, sheetName |
| `calculate_profit_margin` | 计算利润率 | revenueRange, profitRange | fileName, sheetName |

### 数据处理 (ExcelData)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `split_sheet_by_column` | 按列分表 | columnName | fileName, sheetName, dataStartRow |
| `split_and_export` | 分表并导出独立文件 | columnName, outputFolder | fileName, sheetName, fileFormat |
| `merge_sheets` | 合并多个工作表 | 无 | fileName, sheetNames, outputSheetName, includeHeader |
| `merge_workbooks` | 合并多个工作簿 | folderPath | fileName, includeSubfolders, skipEmptySheets |
| `export_sheets` | 批量导出工作表 | outputFolder | fileName, sheetNames, fileFormat |
| `delete_sheets` | 批量删除工作表 | sheetNames | fileName |
| `transpose_columns` | 宽表转长表 | startColumn, fieldName | fileName, sheetName |
| `create_multiple_sheets` | 批量创建工作表 | count | fileName, baseName |
| `generate_payslips` | 生成工资条 | 无 | fileName, sheetName, outputSheetName |

### 数据库 (ExcelDatabase)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `connect_database` | 连接数据库获取表列表 | dbType, connectionString | 无 |
| `execute_query` | 执行SQL查询写入Excel | dbType, connectionString, query | outputFileName, outputSheetName |
| `export_table_to_excel` | 导出数据库表到Excel | dbType, connectionString, tableName | outputFileName, outputSheetName |
| `get_table_structure` | 获取表结构信息 | dbType, connectionString, tableName | 无 |

支持数据库类型：`sqlserver`、`mysql`、`postgresql`、`sqlite`

### REST API (ExcelApi)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `fetch_api_data` | 调用REST API获取数据 | url | method, headers, body, outputFileName, outputSheetName, dataPath |
| `test_api_connection` | 测试API连接 | url | method, headers |

### 增强图表 (ExcelChartEnhance)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `create_word_cloud` | 根据文本列生成词云图片 | textColumn | fileName, sheetName, maxWords, width, height, outputPath |
| `create_dynamic_chart` | 创建动态图表（支持二维表/一维表自动透视） | categoryColumn, valueColumn | fileName, sheetName, chartType, title, tableType(2d/1d), groupColumn, itemColumn, valueColumn1d, skipHeader |
| `create_comparison_chart` | 创建多系列对比图 | categoryColumn, valueColumns | fileName, sheetName, chartType, title |
| `create_pareto_chart` | 创建帕累托图（二八分析） | categoryColumn, valueColumn | fileName, sheetName |
| `create_histogram` | 创建直方图 | valueColumn | fileName, sheetName, binCount |
| `create_box_plot` | 创建箱线图（五数概括） | valueColumns | fileName, sheetName |

词云使用 SkiaSharp 渲染为 PNG 图片；帕累托/直方图/箱线图通过数据预处理+基础图表实现。

### 邮件群发 (ExcelMail)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `configure_smtp` | 配置SMTP服务器 | host, port, username, password | enableSsl, fromAddress |
| `send_email` | 发送单封邮件 | to, subject, body | cc, bcc, isHtml, attachments |
| `batch_send` | 批量发送邮件 | emailColumn, subjectColumn, bodyColumn | fileName, sheetName, ccColumn, isHtml |
| `test_smtp` | 测试SMTP连接 | host, port, username, password | enableSsl |
| `preview_email` | 预览邮件 | to, subject, body | cc, isHtml |

### 文件操作 (ExcelFile)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `list_files` | 列出文件信息 | folderPath | pattern, includeSubfolders |
| `batch_rename` | 批量重命名文件 | oldNameColumn, newNameColumn | folderPath, fileName, sheetName |
| `batch_copy` | 批量复制文件 | fileNameColumn, targetFolder | sourceFolder, fileName, sheetName |
| `batch_move` | 批量移动文件 | fileNameColumn, targetFolder | sourceFolder, fileName, sheetName |
| `batch_delete` | 批量删除文件 | fileNameColumn | folderPath, fileName, sheetName |
| `create_folder` | 创建文件夹 | folderPath | 无 |
| `get_file_info` | 获取文件详细信息 | filePath | 无 |
| `open_folder` | 打开文件夹 | folderPath | 无 |

### 二维码/条形码 (ExcelQR)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `generate_qr_code` | 为指定列数据生成二维码图片 | columnNames, outputFolder | fileName, sheetName, size |
| `generate_barcode` | 为指定列数据生成条形码图片（Code128） | columnName, outputFolder | fileName, sheetName, width, height |
| `scan_qr_code` | 扫描图片中的二维码返回内容 | imagePaths | outputFileName, outputSheetName |
| `scan_qr_code_folder` | 批量扫描文件夹中所有图片的二维码 | folderPath | includeSubfolders, outputFileName, outputSheetName |

使用 ZXing.Net + SkiaSharp 实现，无需安装额外运行时。`columnNames` 为 JSON 数组格式如 `["网址","编号"]`。

### 发票识别 (ExcelInvoice)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `import_xml_invoice` | 导入XML电子发票到Excel | xmlPath | outputFileName, outputSheetName |
| `batch_import_invoices` | 批量导入文件夹中XML发票 | folderPath | outputFileName, outputSheetName, includeSubfolders |
| `get_invoice_fields` | 获取发票可提取字段列表 | 无 | 无 |
| `export_invoice_summary` | 导出发票汇总表 | outputFileName | outputSheetName |
| `ocr_invoice` | OCR识别发票图片/PDF | imagePath | outputFileName, outputSheetName |

**OCR 发票识别说明**：
- 自动检测系统是否安装 PaddleOCR（尝试 `paddleocr --version` 或 `python -c "import paddleocr"`）
- 已安装：调用 PaddleOCR 进行 OCR，用正则提取发票号码、日期、买卖方、金额等字段，写入 Excel
- 未安装：返回详细安装指南（pip/conda/国内镜像三种方式）
- 支持图片（png/jpg/bmp）和 PDF 格式
- 安装命令：`pip install paddlepaddle paddleocr`，PDF 还需 `pip install pymupdf`

### 正则提取 (ExcelRegex)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `extract_by_regex` | 正则提取内容到新列 | columnName | fileName, sheetName, patternType, pattern |
| `get_regex_patterns` | 获取预定义正则模式列表 | 无 | 无 |
| `validate_regex` | 验证正则表达式 | pattern | 无 |

### 目录页 (ExcelToc)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `create_toc_sheet` | 创建目录表 | 无 | fileName, tocSheetName, includeHiddenSheets |
| `create_sheets_from_toc` | 根据目录建表 | linkColumnName | fileName, createSheets |
| `update_toc_hyperlinks` | 更新目录超链接 | columnName | fileName |

### 定时任务 (ExcelSchedule)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `create_task` | 创建定时任务 | taskName, cronExpression, toolName | arguments, description |
| `list_tasks` | 列出所有任务 | 无 | 无 |
| `delete_task` | 删除任务 | taskName | 无 |
| `enable_task` | 启用任务 | taskName | 无 |
| `disable_task` | 禁用任务 | taskName | 无 |
| `run_task` | 立即执行任务 | taskName | 无 |

### Word 文档生成 (DocumentGeneration)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `generate_documents` | 批量生成Word文档 | templatePath, outputFolder | fileName, sheetName, nameColumn, format |
| `preview_document` | 预览文档效果 | templatePath | fileName, sheetName |

### PDF 导出 (ExcelPdf)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `export_sheet_to_pdf` | 导出工作表为PDF | pdfPath | fileName, sheetName |
| `export_workbook_to_pdf` | 导出整个工作簿为PDF | pdfPath | fileName |
| `batch_export_to_pdf` | 批量导出多个工作表为PDF | outputFolder | fileName, sheetNames |
| `export_range_to_pdf` | 导出指定区域为PDF | pdfPath, rangeAddress | fileName, sheetName |

使用 QuestPDF（Community License）实现，无需安装 Excel 或打印机驱动。导出为 A4 横向带表格边框的 PDF。

## 通用参数说明

### 可选参数

大多数工具支持以下可选参数，不指定时使用当前活跃的工作簿/工作表：

| 参数 | 类型 | 说明 |
|------|------|------|
| fileName | string | 工作簿文件名 |
| sheetName | string | 工作表名称 |

### 数据格式约定

- **行列号**：从 1 开始，与 Excel 一致
- **范围地址**：标准 Excel 格式，如 `A1:D10`、`B:B`
- **颜色**：支持中文名称（红色、蓝色）和十六进制（#FF0000）
- **批量数据**：JSON 二维数组，如 `[["姓名","年龄"],["张三",25]]`
- **图表类型**：column / line / pie / bar / area / scatter
- **数据库类型**：sqlserver / mysql / postgresql / sqlite

## 内置 AI 对话集成

TableMagic VSTO 插件的 AI 对话窗体（Form7）已支持调用本 CLI：

1. 优先使用内置 SkillManager（COM 模式，功能最全）
2. 内置工具不可用时自动回退到 CLI 调用
3. CLI 使用 ClosedXML 实现，无需安装 Excel

## 架构

```
TableMagic.Cli/
├── Program.cs                     # CLI入口（命令行 + MCP服务器）
├── skill.md                       # Skill描述文件（Agent发现/安装）
├── mcp-config.json                # MCP配置示例
├── Mcp/
│   ├── McpProtocol.cs             # MCP协议类型定义
│   ├── McpServer.cs               # MCP服务器实现（含参数引导）
│   └── StdioTransport.cs          # stdio传输层
├── Excel/
│   ├── IExcelProvider.cs          # Excel操作抽象接口
│   ├── ClosedXmlExcelProvider.cs  # ClosedXML实现（无需Excel）
│   └── PdfExporter.cs            # QuestPDF PDF导出实现
└── Skills/
    ├── ISkill.cs                  # 通用Skill接口 + SkillResult（含MissingParams）
    ├── SkillManager.cs            # 技能管理器（含参数校验+引导提示）
    ├── ExcelBaseSkill.cs          # 基础操作
    ├── ExcelWorkbookSkill.cs      # 工作簿操作
    ├── ExcelSheetSkill.cs         # 工作表操作
    ├── ExcelCellSkill.cs          # 单元格操作
    ├── ExcelRangeSkill.cs         # 区域操作
    ├── ExcelFormatSkill.cs        # 格式设置
    ├── ExcelChartSkill.cs         # 图表操作
    ├── ExcelPivotSkill.cs         # 数据透视表
    ├── ExcelAnalysisSkill.cs      # 数据分析
    ├── ExcelFinanceSkill.cs       # 财务分析
    ├── ExcelDataSkill.cs          # 数据处理（分表/并表/导删/转置/工资条）
    ├── ExcelDatabaseSkill.cs      # 数据库连接/查询/导出
    ├── ExcelApiSkill.cs           # REST API数据获取
    ├── ExcelChartEnhanceSkill.cs  # 增强图表（词云/帕累托/直方图/箱线图）
    ├── ExcelMailSkill.cs          # 邮件群发
    ├── ExcelFileSkill.cs          # 文件批量操作
    ├── ExcelQRSkill.cs            # 二维码/条形码（ZXing.Net + SkiaSharp）
    ├── ExcelInvoiceSkill.cs       # XML发票导入 + OCR识别（PaddleOCR）
    ├── ExcelRegexSkill.cs         # 正则提取/替换/验证
    ├── ExcelTocSkill.cs           # 目录页生成
    ├── ExcelScheduleSkill.cs      # 定时任务
    ├── DocumentGenerationSkill.cs # Word文档生成
    └── ExcelPdfSkill.cs           # PDF导出（QuestPDF）
```

## 依赖

| 包 | 用途 |
|----|------|
| ClosedXML | Excel文件读写（无需安装Excel） |
| QuestPDF | PDF导出（Community License） |
| ZXing.Net + ZXing.Net.Bindings.SkiaSharp | 二维码/条形码生成与识别 |
| SkiaSharp | 图片渲染（词云/二维码） |
| Microsoft.Data.SqlClient | SQL Server连接 |
| MySql.Data | MySQL连接 |
| Npgsql | PostgreSQL连接 |
| System.Data.SQLite.Core | SQLite连接 |
| PaddleOCR（外部） | OCR发票识别（需单独安装，未安装时自动提示安装方法） |
