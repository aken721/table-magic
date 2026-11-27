# 万象插件 TableMagic v2.5.2.1 发行说明

**发布日期**：2026-11-27
**版本**：v2.5.2.1（VSTO 插件 / CLI / Skill 技能包版本号全部统一）

---

## 🎯 本版亮点

这是一个以**稳定性修复 + 功能对齐**为主的维护版本：彻底修复了 PostgreSQL 连接"静默失败"和数据库选项卡切换导致 Excel 崩溃的问题；动态图新增一维表支持；底层依赖升级修复安全漏洞；CLI / Skill 技能包与插件版功能对齐。

---

## ✨ 新增 | Added

- **动态图支持一维表**
  - 图表增强新增"表格类型"选择：二维表 / 一维表
  - 二维表：沿用原有"列标题 / 行标题"分组方式
  - 一维表：自动从选区首行 / 首列加载"分组 / 类别 / 数值"三个选项，将一维明细表自动透视成二维数据后再生成动态图，支持"是否跳过标题行"
  - 示例：gdp.xlsx 选中 D1:F900，分组选 `period`、类别选 `region_name`、数值选 `value`，即可生成按年份分组、各省 GDP 变化的动态图
  - CLI / Skill 的 `create_dynamic_chart` 工具同步支持 `tableType=1d`、`groupColumn`、`itemColumn`、`valueColumn1d`、`skipHeader` 参数
- **标题输入框占位符**：图表标题、X 轴标题、Y 轴标题为空时显示浅灰色"（可不填）"，获得焦点自动清空
- **依赖升级（安全修复）**：DynamicPic.Core 升级至 1.0.3，Magick.NET 升级至 15.1.0 修复已知安全漏洞；同时解耦 Excel 互操作 PIA 版本（改用 late-bound COM 调用，不再硬绑定 PIA 版本）

## 🔧 优化 | Changed

- **获取选区按钮**：选区仅为单个单元格时，自动改用当前工作表的 UsedRange 作为数据范围，避免只取一个格子
- **图表增强性能**：注释掉未启用的系统字体枚举代码，消除切换标签时的界面卡顿（代码保留，取消注释即可恢复字体选择功能）
- **安装包改进**：改用 bin\Release 构建；移除 CASPOL 自定义动作与 macOS 动态库；新增 .NET Framework 4.8 与 VSTO 4.0 Runtime 先决条件（自动检测并按需要下载安装运行时）；关闭无效的 Authenticode 签名，产物为无签名的干净安装包
- **CLI / Skill 版本与文档**：版本号统一为 2.5.2.1；修正 `create_dynamic_chart` / `connect_database` 工具参数文档与实际实现不一致的问题
- **数据库错误统一处理**：所有数据库操作统一错误提示，并写入日志 `%LOCALAPPDATA%\TableMagic\logs\database.log`，便于排查问题

## 🐛 修复 | Fixed

- **修复 PostgreSQL 连接失败无任何提示**：统一异常上报并显示真实错误原因；支持非 public schema（显示为 `schema.table`）；用双引号正确包裹标识符（兼容大写、中文、含空格的表名）；使用 `NpgsqlConnectionStringBuilder` 正确构建连接串
- **修复数据库选项卡切换导致 Excel 整体崩溃**：ComboBox 在 DataSource 绑定状态下调用 `Items.Clear()` 抛出 `ArgumentException`，且 `async void` 事件处理器无法捕获，最终未处理异常终结 Excel；现改为先解绑再清空，并为 async void 事件处理器、WinForms UI 线程和 AppDomain 添加全局异常兜底，个别异常不再拖垮 Excel
- **修复定时器事件重复订阅**：`StartTimer()` 每次调用累积一个处理器，现先取消再订阅，并为定时器回调加判空和异常保护
- **CLI / Skill 同步修复 PostgreSQL**：CLI 与插件内两份 Skill 副本均应用上述 PostgreSQL 修复与一维表动态图逻辑，与 VSTO 版功能对齐

---

## 📥 下载与安装

### Excel / WPS 插件

| 文件 | 说明 |
|---|---|
| `TableMagic_x86.exe` | 单文件安装包（内含 MSI + VSTO 4.0 Runtime 先决条件），双击安装即可；会自动检测并安装 VSTO 运行时 |
| `InvoiceOCR.7z` | PaddleOCR 发票识别模型（可选，需要发票 OCR 功能时解压到安装目录）|

1. 从 [GitHub Releases](https://github.com/aken721/table-magic/releases) 或 [GitCode Releases](https://gitcode.com/aken721/table-magic/releases) 下载 `TableMagic_x86.exe`
2. 双击安装，安装完成后打开 Excel，功能区自动出现 **"万象插件"** 选项卡
3. 安装包已配置 WPS 注册表支持项，WPS 中可直接使用
4. 若需发票 OCR，将 `InvoiceOCR.7z` 解压到安装目录

> 注：如需不含 OCR 模型的精简包，可自行用 `TableMagic_noPaddle.aip` 通过 Advanced Installer 构建。

### CLI / Skill 技能包（Agent / MCP 用）

| 文件 | 说明 |
|---|---|
| `table-magic-skill.zip` | 预编译单文件技能包（自包含 win-x64，约 49.5 MB），内含 `tablemagic.exe`、`skill.md`、`mcp-config.json`、`README.md` |
| `TableMagic.Cli.2.5.2.1.nupkg` | dotnet tool 包（若仓库提供）|

- **ZIP 安装（推荐）**：解压 `table-magic-skill.zip` 到 Agent 技能安装目录，Agent 会自动通过 `skill.md` 和 `mcp-config.json` 发现并注册技能；也可把 zip 拖入 Agent 对话框让其自动安装
- **dotnet tool 安装**：`dotnet tool install --global --add-source ./nupkg TableMagic.Cli`
- MCP 配置：
  ```json
  { "mcpServers": { "tablemagic": { "command": "<path>/tablemagic", "args": ["mcp"] } } }
  ```

---

## 🔧 系统要求

| 项目 | 要求 |
|---|---|
| 操作系统 | Windows 7 及以上（推荐 Windows 10 / 11）|
| Office | Microsoft Office 2013 及以上，或 WPS Office；VSTO 4.0 Runtime 由安装包自动处理 |
| .NET Framework | 4.8（安装包内含启动条件检测）|
| CLI / Skill | .NET 8.0（ZIP 版为自包含，无需单独安装）|

---

## ⚠️ 注意事项

- 本版安装包**不再进行 Authenticode 数字签名**（此前为无效的自签名测试证书），首次安装如弹出 SmartScreen "未知发布者" 提示，选择"更多信息 → 仍要运行"即可。
- 字体选择功能当前未启用（词云图使用默认字体），相关代码保留，后续可随时恢复。
- 数据库相关异常请查看日志 `%LOCALAPPDATA%\TableMagic\logs\database.log`。

---

**完整更新日志**：[CHANGELOG.md](./CHANGELOG.md)
