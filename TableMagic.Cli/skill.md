# TableMagic Skill

## 元数据

- **名称**: tablemagic
- **版本**: 2.5.2.1
- **描述**: Excel 操作技能，提供工作簿/工作表/单元格/区域/格式/图表/分析/数据库/API/邮件/文件/二维码/发票/正则/目录/定时/PDF等 111 个工具
- **协议**: MCP (Model Context Protocol) 2024-11-05
- **传输**: stdio
- **运行时**: .NET 8.0
- **依赖**: 无需安装 Excel（使用 ClosedXML）

## 安装

### 从 ZIP 安装（推荐，无需 .NET SDK）

1. 将 `table-magic-skill.zip` 解压到任意目录（称为**技能安装目录**）
2. 确认安装目录中包含 `tablemagic.exe`
3. 在 Agent 的 MCP 配置中，`command` 字段填写**技能安装目录中的 tablemagic.exe 完整路径**，或确保该目录在系统 PATH 中

### 前置条件

- ZIP 方式：无需任何前置条件（self-contained，包含 .NET 运行时）
- 源码/工具方式：需要 .NET 8.0 SDK 或运行时

### 从源码构建

```bash
git clone <repo-url>
cd TableMagic
dotnet build TableMagic.Cli
dotnet run --project TableMagic.Cli -- mcp
```

### 发布为独立可执行文件

```bash
dotnet publish TableMagic.Cli -c Release -r win-x64 --self-contained -o ./publish
./publish/tablemagic mcp
```

### 全局工具安装

```bash
dotnet pack TableMagic.Cli
dotnet tool install --global --add-source ./nupkg TableMagic.Cli
tablemagic mcp
```

## Agent 配置

在 Agent 的 MCP 配置文件中添加以下内容。**重要**：`command` 必须指向技能安装目录中的 `tablemagic.exe`，不要使用开发环境路径。

### ZIP 安装方式（推荐）

将 `table-magic-skill.zip` 解压后，用解压目录中的完整路径配置：

```json
{
  "mcpServers": {
    "tablemagic": {
      "command": "<技能安装目录>/tablemagic.exe",
      "args": ["mcp"]
    }
  }
}
```

例如解压到 `C:\Skills\tablemagic`，则配置为：

```json
{
  "mcpServers": {
    "tablemagic": {
      "command": "C:/Skills/tablemagic/tablemagic.exe",
      "args": ["mcp"]
    }
  }
}
```

### 全局工具安装方式

```json
{
  "mcpServers": {
    "tablemagic": {
      "command": "tablemagic",
      "args": ["mcp"]
    }
  }
}
```

### 源码运行方式

```json
{
  "mcpServers": {
    "tablemagic": {
      "command": "dotnet",
      "args": ["run", "--project", "<path-to>/TableMagic.Cli", "mcp"]
    }
  }
}
```

### 各 Agent 配置位置

| Agent | 配置文件路径 |
|-------|-------------|
| OpenClaw | `~/.openclaw/mcp.json` |
| Hermes | `~/.hermes/mcp_config.json` |
| WorkBuddy | `~/.workbuddy/mcp.json` |
| QClaw | 项目根目录 `qclaw.mcp.json` |
| Trae Work | 项目根目录 `.trae/mcp.json` |

## 行为约束

### 可执行文件路径（重要）

**禁止**使用开发环境路径（如 `E:\SourceCode\Csharp\TableMagic\TableMagic.Cli\bin\...`）调用 tablemagic.exe。
**必须**使用技能安装目录中的 tablemagic.exe。技能安装目录通常位于：
- `~/.qoderworkcn/skills/table-magic/`
- `~/.openclaw/skills/table-magic/`
- 或其他 Agent 的 skills 子目录

正确的调用方式：
```
C:\Users\Ken\.qoderworkcn\skills\table-magic\tablemagic.exe mcp
```

错误的调用方式（禁止）：
```
E:\SourceCode\Csharp\TableMagic\TableMagic.Cli\bin\Release\net8.0\tablemagic.exe mcp
```

### 参数引导

当用户 prompt 不明确或缺少必需参数时，**必须**向用户追问缺失的参数，不要猜测或使用默认值。具体规则：

1. **必需参数缺失**：返回错误信息，明确列出缺失的参数名称和描述，引导用户补充
2. **参数值模糊**：如用户说"创建一个表格"但未指定文件名，应追问"请提供工作簿文件名"
3. **参数类型错误**：如行号传入了字符串，应提示"row 参数应为整数"
4. **一次只追问最关键的缺失参数**，避免一次列出过多问题让用户困惑

### 性能优化（重要）

**必须**使用以下策略以获得最佳性能：

1. **优先使用复合工具**，避免多次单独调用：
   - 创建工作簿+多工作表 → 用 `create_workbook_with_sheets`（一次调用），而非 create_workbook + N次 create_worksheet
   - 写入完整表格 → 用 `write_table_data`（一次调用），而非 N次 set_cell_value
   - 读取完整表格 → 用 `read_table_data`（一次调用），而非 N次 get_cell_value
   - 批量执行不同工具 → 用 `batch_call`（一次调用），而非多次 tools/call

2. **禁止**使用 `tablemagic.exe call` 命令行方式逐个调用工具，这会为每次调用启动新进程，极其缓慢
3. **必须**通过 MCP 协议（`tools/call`）调用工具，MCP 模式下单进程持久运行，无启动开销

### 错误处理

- 工具执行失败时，返回结构化错误信息，包含：错误描述、处理建议、是否需要用户决策
- Agent 收到错误后应向用户说明情况，并询问是否尝试其他方式

## 工具清单

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

## 参数说明

### 通用可选参数

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
| `connect_database` | 连接数据库获取表列表 | dbType, connectionString | host, port, database, user, password（PostgreSQL 分散字段，优先于 connectionString） |
| `execute_query` | 执行SQL查询写入Excel | dbType, connectionString, query | outputFileName, outputSheetName |
| `export_table_to_excel` | 导出数据库表到Excel | dbType, connectionString, tableName | outputFileName, outputSheetName |
| `get_table_structure` | 获取表结构信息 | dbType, connectionString, tableName | 无 |

### API接口 (ExcelApi)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `fetch_api_data` | 调用REST API获取数据 | url | method, headers, body, outputFileName, outputSheetName, dataPath |
| `test_api_connection` | 测试API连接 | url | method, headers |

### 增强图表 (ExcelChartEnhance)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `create_word_cloud` | 生成词云 | textColumn | fileName, sheetName, maxWords, width, height |
| `create_dynamic_chart` | 创建动态图表（支持二维表/一维表，一维表自动透视成二维表） | categoryColumn, valueColumn（二维表）；一维表另需 groupColumn, itemColumn, valueColumn1d | fileName, sheetName, chartType, title, tableType(2d/1d), skipHeader |
| `create_comparison_chart` | 创建对比图 | dataRange | fileName, sheetName, title |
| `create_pareto_chart` | 创建帕累托图 | dataRange, valueColumn, categoryColumn | fileName, sheetName |
| `create_histogram` | 创建直方图 | dataRange | fileName, sheetName, binCount |
| `create_box_plot` | 创建箱线图 | dataRange | fileName, sheetName |

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
| `generate_qr_code` | 生成二维码 | columnNames, outputFolder | fileName, sheetName, size |
| `generate_barcode` | 生成条形码 | columnName, outputFolder | fileName, sheetName, width, height |
| `scan_qr_code` | 扫描二维码 | imagePaths | outputSheetName |
| `scan_qr_code_folder` | 批量扫描二维码 | folderPath | includeSubfolders, outputSheetName |

### 发票识别 (ExcelInvoice)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `import_xml_invoice` | 导入XML发票 | xmlPath | outputFileName, outputSheetName |
| `batch_import_invoices` | 批量导入XML发票 | folderPath | outputFileName, outputSheetName, includeSubfolders |
| `get_invoice_fields` | 获取发票字段列表 | 无 | 无 |
| `export_invoice_summary` | 导出发票汇总 | outputFileName | outputSheetName |
| `ocr_invoice` | OCR识别发票图片/PDF（自动检测PaddleOCR，未安装时返回安装指南） | imagePath | outputFileName, outputSheetName |

### 正则提取 (ExcelRegex)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `extract_by_regex` | 正则提取内容到新列 | columnName | fileName, sheetName, patternType, pattern |
| `get_regex_patterns` | 获取预定义正则模式 | 无 | 无 |
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

### Word文档生成 (DocumentGeneration)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `generate_documents` | 批量生成Word文档 | templatePath, outputFolder | fileName, sheetName, nameColumn, format |
| `preview_document` | 预览文档效果 | templatePath | fileName, sheetName |

### PDF导出 (ExcelPdf)

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|---------|---------|
| `export_sheet_to_pdf` | 导出工作表为PDF | pdfPath | fileName, sheetName |
| `export_workbook_to_pdf` | 导出工作簿为PDF | pdfPath | fileName |
| `batch_export_to_pdf` | 批量导出为PDF | outputFolder | fileName, sheetNames |
| `export_range_to_pdf` | 导出区域为PDF | pdfPath, rangeAddress | fileName, sheetName |

## 使用示例

```
用户: 创建一个销售数据表
Agent: 调用 create_workbook(fileName: "销售数据.xlsx", sheetName: "销售")

用户: 在A1写入标题
Agent: 调用 set_cell_value(fileName: "销售数据.xlsx", sheetName: "销售", row: 1, column: 1, value: "销售记录")

用户: 给标题加粗居中
Agent: 调用 set_cell_format(rangeAddress: "A1", bold: true, horizontalAlignment: "center")

用户: 分析B列的数据
Agent: 调用 get_range_statistics(range: "B:B")
```

## CLI 直接调用

```bash
# 列出工具
tablemagic list-tools

# 创建工作簿
tablemagic call create_workbook --args '{"fileName":"test.xlsx"}'

# 写入数据
tablemagic call set_cell_value --args '{"fileName":"test.xlsx","row":1,"column":1,"value":"Hello"}'

# 读取数据
tablemagic call get_range_values --args '{"fileName":"test.xlsx","rangeAddress":"A1:D10"}'

# 批量执行
tablemagic batch calls.json
```