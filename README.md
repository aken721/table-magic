<div align="center">

# 🧰 万象插件 | TableMagic Add-in

**装上它，Excel帮你摸鱼**

[🇬🇧 English](./README_en.md) · [![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT) · [![Version](https://img.shields.io/badge/Version-v2.5.1.0-blue.svg)](https://gitee.com/aken721/table-magic/releases) · [![Platform](https://img.shields.io/badge/Platform-Office%20%2B%20WPS-orange.svg)]() · [![Stars](https://img.shields.io/badge/Stars-⭐%20Give%20a%20Star-yellow.svg)](https://gitee.com/aken721/table-magic)

[📥 立即下载](https://gitee.com/aken721/table-magic/releases) · [🐛 提交问题](https://gitee.com/aken721/table-magic/issues) · [📧 定制开发](mailto:aken721@163.com)

</div>

---

### 简介

**万象插件（TableMagic）**是一款完全**免费开源**的Excel增强插件，基于VSTO技术开发，兼容微软Office和WPS。

不同于普通工具箱只提供基础表格操作，本插件聚焦于**真实工作场景中的高频痛点**——将发票处理、邮件群发、合同批量生成、定时自动化、AI自然语言操作等复杂任务，封装为**几步点击即可完成的交互式操作**，让没有技术背景的普通打工人也能享受自动化带来的效率提升。

### ✨ 核心功能

#### 📊 表格增强

**数据拆分与汇总**
- **分表/分簿**：按指定列条件一键拆分数据，生成独立工作表或独立工作簿，支持保留标题行
- **合并表/合并工作簿**：多Sheet、多工作簿数据快速纵向汇总，跨文件整合不再繁琐
- **从文件夹批量导入**：将整个文件夹内所有工作簿的指定表批量合并到当前工作簿

**表结构转换**
- **二维表转一维表**（Unpivot）：将交叉汇总表还原为明细数据，日期列自动格式化、数字列自动去除千位符，是数据透视/分析的必备前处理工具
- **工资表→工资条**：一键将汇总式工资表拆分为每人独立工资条，可设置间隔行数

**数据提取与清洗**
- **正则提取数据**：内置9种预设规则（数字 / 纯英文 / 中文 / 网址 / 身份证号 / 邮箱 / 电话 / IP地址 / 自定义正则），从单元格中精准提取特定格式内容
- **序号列智能重排**：分表/并表操作后自动检测并重新为"序号"列连续编号

**工作表管理**
- **Excel目录页**（双向功能）：
  - 从现有所有工作表自动生成带超链接的目录索引页，一键导航任意Sheet
  - 从目录清单批量创建对应的空白工作表，快速搭建多Sheet报告框架
- **批量建空白表**：1-15张批量创建，支持自定义前缀命名
- **批量删表**：按规则批量删除工作表
- **聚光灯**：点击单元格时高亮整行整列，支持6种主题色切换，告别"找行找列"烦恼
- **Excel转PDF**：支持当前表、全部表转单PDF、全部表转多PDF，可设置页面大小和缩放

#### 🧾 发票管理（独有功能）
- **XML格式数电票**：批量读取电子发票主要信息到Excel
- **PDF/图片发票OCR识别**：基于PaddleOCR，批量识别扫描件/照片中的发票数据
- 财务报销从此告别手工录入！

#### 📝 Word模板批量生成（独有功能）
- 在Word模板中设置占位符，Excel每行数据自动生成一份独立文档
- 支持**图片插入**（在Excel中填写图片路径，自动嵌入Word对应位置）
- 支持 `.doc` 和 `.docx` 格式自动适配
- 典型场景：批量合同、录用通知、员工证书、物料标签

#### 📧 邮件群发（独有功能）
- 基于Excel表格，对不同收件人发送**个性化内容和附件**
- 内置**HTML邮件编辑器**（自研富文本控件），支持图片、排版，发送精美HTML邮件
- 再也不用逐一复制粘贴！

#### ⏰ 定时任务（独有功能）
- Windows任务计划风格的可视化定时设置
- 支持执行：**CMD命令 / VBA宏代码 / BAT脚本 / Python脚本 / EXE程序**
- 触发器：一次性、每天、每周、每月、自定义间隔
- 无需离开Excel即可配置并持久化任务计划

#### 📁 批量文件管理
- **批量重命名**：将目录文件读入Excel，在表格中灵活编辑新名称后批量执行
- **批量移动/删除**：支持包含子目录
- 路径和文件名支持超链接，点击直接打开

#### 📈 图表增强
- **HTML交互式动态图表**：生成可交互的网页图表（基于ECharts），支持缩放、筛选、动态展示
- **GIF动图图表**：生成动态演示图，嵌入PPT或报告立即吸引眼球
- **帕累托图**：自动计算累计占比，一键生成柱+线复合帕累托分析图
- **词云图**：统计单元格词频，一键生成词云，支持导出为SVG矢量图或PNG位图

#### 🔲 二维码 & 条形码
- 生成**彩色二维码**（自定义前景色/背景色，支持嵌入Logo图片）
- **多列合并生成**：按 `key:value;` 格式将多列数据拼接后批量生成二维码
- 生成 **Code128条形码**
- 识别图片或摄像头中的二维码/条形码，提取数据到Excel

#### 🎵 背景音乐
- 在Excel中播放背景音乐，告别枯燥数据录入
- 支持 mp3 / wav / flac / aiff / wma / aac / mp4 等主流格式

#### 🗄️ 数据库直连导出（独有功能）
- 支持6种主流数据库：**MySQL / SQL Server / Access / SQLite / PostgreSQL / Oracle**
- 关键字模糊搜索表名，DataGridView实时预览数据
- 长数字自动转文本格式，避免科学计数法截断问题
- Access / SQLite 支持双击直接选择本地文件
- 数据库表格一键导出到Excel，无需写SQL、无需编程

#### 🌐 API接口数据写入（独有功能）
- 填写接口URL，数据自动写入当前Excel表格，无需Python/VBA
- 支持3种认证方式：**无认证 / 用户名+密码（Basic Auth）/ API Key**（支持自定义Header名称与值前缀）
- 支持URL占位符替换（动态参数化接口地址）
- 自动适配 TLS 1.0 / 1.1 / 1.2，兼容老旧服务端
- 典型场景：对接企业内网系统、气象/财经/政务数据抓取

#### 🤖 AI对话 & 自然语言操作（v2.5重大更新）
- 支持 DeepSeek、Qwen（通义千问）、Kimi、GLM、ChatGPT、Claude 等14种主流API
- 支持 Ollama / LM Studio / vLLM / llama.cpp 等**本地模型**（数据不出本机）
- 自动识别模型能力，智能切换 Function Calling / 提示工程模式
- **Skills插件化扩展架构**：通过标准 ISkill 接口加载第三方扩展，覆盖工作簿、工作表、单元格、区域、格式、图表、数据透视表、财务分析等24种内置技能
- 对话即操作：用自然语言描述需求，AI自动完成Excel操作

### 📥 安装方式

1. 前往 [Release页面](https://gitee.com/aken721/table-magic/releases) 下载 `.msi` 安装包
2. v2.4.6.2起需同时下载 `InvoiceOCR.7z`（OCR模型），解压到插件安装目录
3. 完整包可从 [GitHub](https://github.com/aken721/table-magic/releases) 或 [GitCode](https://gitcode.com/aken721/table-magic/releases) 下载
4. 安装后打开Excel，功能区自动出现 **"智能工具箱"** 选项卡（后续版本将更名为"万象插件"）

### 🔧 系统要求

| 项目 | 要求 |
|:---|:---|
| 操作系统 | Windows 7 及以上 |
| Office | Microsoft Office 2013+ 或 WPS |
| .NET | .NET Framework 4.7.2+ |
| 发票OCR | 需额外下载OCR模型包 |

### 💡 适用人群

- 📊 **财务/会计**：发票批量录入、财务报表自动化
- 👥 **HR/行政**：合同批量生成、通知群发、文件重命名
- 📈 **运营/销售**：数据汇总、报告自动化
- 🏭 **采购/供应链**：数据库导入、价格批量处理
- 任何在Excel上花费大量重复时间的打工人

### 🤝 定制开发

每个企业的业务场景都不同，现有功能或许无法完全满足你的需求。欢迎联系进行**个性化定制开发**：

📧 **aken721@163.com**

### 📜 开源协议

MIT License — 完全免费，可商用，保留版权声明即可。

Copyright (c) 2023 Ken He / 海南纽沃数字科技有限公司

---

<div align="center">

**如果这个项目对你有帮助，请给个 ⭐ Star 支持一下！**

[Gitee](https://gitee.com/aken721/table-magic) · [GitHub](https://github.com/aken721/table-magic) · [GitCode](https://gitcode.com/akin721/table-magic)

📋 [查看完整版本更新日志 →](./CHANGELOG.md)

</div>
