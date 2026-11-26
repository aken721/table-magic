using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelChartEnhanceSkill : ISkill
    {
        public string Name => "ExcelChartEnhance";
        public string Description => "图表增强技能，支持词云图、动态图表等高级图表功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "create_word_cloud",
                    Description = "根据数据创建词云图。当用户要求生成词云、词云图时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "textColumn", new { type = "string", description = "文本数据列名" } },
                                { "weightColumn", new { type = "string", description = "权重列名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "maxWords", new { type = "integer", description = "最大词数（默认100）" } },
                                { "outputPath", new { type = "string", description = "输出图片路径（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "textColumn" }
                },
                new SkillTool
                {
                    Name = "create_dynamic_chart",
                    Description = "创建动态图表，支持按时间或类别动态展示数据变化。当用户要求动态图表、动画图表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "categoryColumn", new { type = "string", description = "类别列名" } },
                                { "valueColumn", new { type = "string", description = "数值列名" } },
                                { "timeColumn", new { type = "string", description = "时间列名（可选，用于动态展示）" } },
                                { "chartType", new { type = "string", description = "图表类型：bar/line/pie/scatter，默认bar" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "tableType", new { type = "string", description = "表格类型：2d(默认)/1d(一维表自动透视成二维表)" } },
                                { "groupColumn", new { type = "string", description = "一维表分组列（如 period，按年份分组）" } },
                                { "itemColumn", new { type = "string", description = "一维表类别列（如 region_name，各省）" } },
                                { "valueColumn1d", new { type = "string", description = "一维表数值列（如 value，GDP）" } },
                                { "skipHeader", new { type = "string", description = "是否跳过首行表头，默认true" } },
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "categoryColumn", "valueColumn" }
                },
                new SkillTool
                {
                    Name = "create_comparison_chart",
                    Description = "创建对比图表，支持多系列数据对比。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "categoryColumn", new { type = "string", description = "类别列名" } },
                                { "valueColumns", new { type = "string", description = "数值列名列表（JSON数组格式）" } },
                                { "chartType", new { type = "string", description = "图表类型：bar/line/radar，默认bar" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "categoryColumn", "valueColumns" }
                },
                new SkillTool
                {
                    Name = "create_pareto_chart",
                    Description = "创建帕累托图（帕累托分析）。当用户要求帕累托图、二八分析时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "categoryColumn", new { type = "string", description = "类别列名" } },
                                { "valueColumn", new { type = "string", description = "数值列名" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "categoryColumn", "valueColumn" }
                },
                new SkillTool
                {
                    Name = "create_histogram",
                    Description = "创建直方图。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "valueColumn", new { type = "string", description = "数值列名" } },
                                { "binCount", new { type = "integer", description = "分组数量（默认10）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "valueColumn" }
                },
                new SkillTool
                {
                    Name = "create_box_plot",
                    Description = "创建箱线图。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "valueColumns", new { type = "string", description = "数值列名列表（JSON数组格式）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "valueColumns" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "create_word_cloud":
                        return await CreateWordCloudAsync(arguments);
                    case "create_dynamic_chart":
                        return await CreateDynamicChartAsync(arguments);
                    case "create_comparison_chart":
                        return await CreateComparisonChartAsync(arguments);
                    case "create_pareto_chart":
                        return await CreateParetoChartAsync(arguments);
                    case "create_histogram":
                        return await CreateHistogramAsync(arguments);
                    case "create_box_plot":
                        return await CreateBoxPlotAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelChartEnhanceSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private Task<SkillResult> CreateWordCloudAsync(Dictionary<string, object> arguments)
        {
            var textColumn = arguments["textColumn"].ToString();
            var weightColumn = arguments.ContainsKey("weightColumn") ? arguments["weightColumn"].ToString() : null;
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
            var maxWords = arguments.ContainsKey("maxWords") ? Convert.ToInt32(arguments["maxWords"]) : 100;
            var outputPath = arguments.ContainsKey("outputPath") ? arguments["outputPath"].ToString() : null;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

            var usedRange = sheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;

            var columnMap = new Dictionary<string, int>();
            for (int c = 1; c <= lastCol; c++)
            {
                var colName = sheet.Cells[1, c].Text?.ToString();
                if (!string.IsNullOrEmpty(colName))
                    columnMap[colName] = c;
            }

            if (!columnMap.ContainsKey(textColumn))
                return Task.FromResult(new SkillResult { Success = false, Error = $"未找到文本列: {textColumn}" });

            int textColIdx = columnMap[textColumn];
            int weightColIdx = weightColumn != null && columnMap.ContainsKey(weightColumn) ? columnMap[weightColumn] : 0;

            var wordWeights = new Dictionary<string, double>();

            for (int r = 2; r <= lastRow; r++)
            {
                var text = sheet.Cells[r, textColIdx].Text?.ToString();
                if (string.IsNullOrEmpty(text)) continue;

                var weight = weightColIdx > 0 ? GetNumericValue(sheet.Cells[r, weightColIdx]) : 1;

                var words = text.Split(new[] { ' ', ',', '，', '、', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var word in words)
                {
                    var w = word.Trim();
                    if (string.IsNullOrEmpty(w)) continue;

                    if (wordWeights.ContainsKey(w))
                        wordWeights[w] += weight;
                    else
                        wordWeights[w] = weight;
                }
            }

            var topWords = wordWeights.OrderByDescending(x => x.Value).Take(maxWords).ToList();

            var existingNames = new List<string>();
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            string chartSheetName = "词云_" + DateTime.Now.ToString("HHmmss");
            int suffix = 2;
            while (existingNames.Any(n => string.Equals(n, chartSheetName, StringComparison.OrdinalIgnoreCase)))
            {
                chartSheetName = $"词云_{DateTime.Now.ToString("HHmmss")}_{suffix}";
                suffix++;
            }

            var chartSheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            chartSheet.Name = chartSheetName;

            chartSheet.Cells[1, 1].Value = "词语";
            chartSheet.Cells[1, 2].Value = "频次";

            for (int i = 0; i < topWords.Count; i++)
            {
                chartSheet.Cells[i + 2, 1].Value = topWords[i].Key;
                chartSheet.Cells[i + 2, 2].Value = topWords[i].Value;
            }

            var chart = chartSheet.Shapes.AddChart2(
                Style: -1,
                XlChartType: Excel.XlChartType.xlColumnClustered,
                Left: chartSheet.Range["D1"].Left,
                Top: chartSheet.Range["D1"].Top,
                Width: 400,
                Height: 300
            ).Chart;

            chart.SetSourceData(chartSheet.Range[$"A1:B{topWords.Count + 1}"]);
            chart.HasTitle = true;
            chart.ChartTitle.Text = "词频统计";

            chartSheet.Activate();

            return Task.FromResult(new SkillResult { Success = true, Content = $"词云数据已生成，共 {topWords.Count} 个词语\n数据已写入工作表: {chartSheet.Name}" });
        }

        private Task<SkillResult> CreateDynamicChartAsync(Dictionary<string, object> arguments)
        {
            var categoryColumn = arguments["categoryColumn"].ToString();
            var valueColumn = arguments["valueColumn"].ToString();
            var timeColumn = arguments.ContainsKey("timeColumn") ? arguments["timeColumn"].ToString() : null;
            var chartType = arguments.ContainsKey("chartType") ? arguments["chartType"].ToString().ToLower() : "bar";
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
            string tableType = arguments.ContainsKey("tableType") ? arguments["tableType"].ToString().ToLower() : "2d";
            bool oneDimension = tableType == "1d";
            bool skipHeader = !arguments.ContainsKey("skipHeader") || arguments["skipHeader"].ToString() != "false";

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

            var usedRange = sheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;

            var columnMap = new Dictionary<string, int>();
            for (int c = 1; c <= lastCol; c++)
            {
                var colName = sheet.Cells[1, c].Text?.ToString();
                if (!string.IsNullOrEmpty(colName))
                    columnMap[colName] = c;
            }

            Excel.Worksheet chartSheet = sheet;
            string dataRangeText;
            string chartTitleText;

            if (oneDimension)
            {
                // 一维表：按分组列/类别列/数值列读取，透视成二维表后生成图表
                string groupColumn = arguments.ContainsKey("groupColumn") ? arguments["groupColumn"].ToString() : categoryColumn;
                string itemColumn = arguments.ContainsKey("itemColumn") ? arguments["itemColumn"].ToString() : null;
                string valueColumn1d = arguments.ContainsKey("valueColumn1d") ? arguments["valueColumn1d"].ToString() : valueColumn;
                if (string.IsNullOrEmpty(itemColumn) || !columnMap.ContainsKey(groupColumn) ||
                    !columnMap.ContainsKey(itemColumn) || !columnMap.ContainsKey(valueColumn1d))
                    return Task.FromResult(new SkillResult { Success = false, Error = "一维表模式需要有效的 groupColumn/itemColumn/valueColumn1d" });

                int gIdx = columnMap[groupColumn], iIdx = columnMap[itemColumn], vIdx = columnMap[valueColumn1d];
                int startRow = skipHeader ? 2 : 1;
                var groups = new Dictionary<string, Dictionary<string, decimal>>();
                var groupOrder = new List<string>();
                var itemOrder = new List<string>();
                for (int r = startRow; r <= lastRow; r++)
                {
                    string g = sheet.Cells[r, gIdx].Text?.ToString()?.Trim() ?? "";
                    string it = sheet.Cells[r, iIdx].Text?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(g) || string.IsNullOrEmpty(it)) continue;
                    decimal val;
                    if (!decimal.TryParse(sheet.Cells[r, vIdx].Text?.ToString(), out val)) continue;
                    if (!groups.ContainsKey(g)) { groups[g] = new Dictionary<string, decimal>(); groupOrder.Add(g); }
                    if (!groups[g].ContainsKey(it)) itemOrder.Add(it);
                    groups[g][it] = val;
                }
                if (groups.Count == 0 || itemOrder.Count == 0)
                    return Task.FromResult(new SkillResult { Success = false, Error = "一维表未读取到有效数据" });

                // 新建透视 sheet
                string pivotName = "_动态图透视_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var pivot = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                pivot.Name = pivotName;
                pivot.Cells[1, 1].Value = "类别";
                for (int c = 0; c < groupOrder.Count; c++) pivot.Cells[1, c + 2].Value = groupOrder[c];
                for (int r = 0; r < itemOrder.Count; r++)
                {
                    pivot.Cells[r + 2, 1].Value = itemOrder[r];
                    for (int c = 0; c < groupOrder.Count; c++)
                    {
                        decimal v;
                        pivot.Cells[r + 2, c + 2].Value = groups[groupOrder[c]].TryGetValue(itemOrder[r], out v) ? (object)v : "";
                    }
                }
                chartSheet = pivot;
                dataRangeText = $"A1:{ColLetter(groupOrder.Count + 1)}{itemOrder.Count + 1}";
                chartTitleText = string.IsNullOrEmpty(titleArg(arguments)) ? "动态图表（一维表）" : titleArg(arguments);
            }
            else
            {
                if (!columnMap.ContainsKey(categoryColumn) || !columnMap.ContainsKey(valueColumn))
                    return Task.FromResult(new SkillResult { Success = false, Error = "未找到指定的列" });
                dataRangeText = $"${categoryColumn}$1:${categoryColumn}${lastRow},${valueColumn}$1:${valueColumn}${lastRow}";
                chartTitleText = $"{categoryColumn} - {valueColumn}";
            }

            var xlChartType = chartType switch
            {
                "line" => Excel.XlChartType.xlLine,
                "pie" => Excel.XlChartType.xlPie,
                "scatter" => Excel.XlChartType.xlXYScatter,
                _ => Excel.XlChartType.xlColumnClustered
            };

            var chart = chartSheet.Shapes.AddChart2(
                Style: -1,
                XlChartType: xlChartType,
                Left: chartSheet.Cells[1, lastCol + 2].Left,
                Top: chartSheet.Cells[1, lastCol + 2].Top,
                Width: 500,
                Height: 350
            ).Chart;
            chart.SetSourceData(chartSheet.Range[dataRangeText]);
            chart.HasTitle = true;
            chart.ChartTitle.Text = chartTitleText;

            if (chartType == "bar" || chartType == "line")
            {
                chart.Axes(Excel.XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(Excel.XlAxisType.xlCategory).AxisTitle.Text = categoryColumn;
                chart.Axes(Excel.XlAxisType.xlValue).HasTitle = true;
                chart.Axes(Excel.XlAxisType.xlValue).AxisTitle.Text = valueColumn;
            }

            string mode = oneDimension ? "一维表→二维表" : "二维表";
            return Task.FromResult(new SkillResult { Success = true, Content = $"动态图表已创建（{mode}）\n类型: {chartType}\n标题: {chartTitleText}" });
        }

        private static string titleArg(Dictionary<string, object> arguments)
        {
            return arguments.ContainsKey("title") ? arguments["title"].ToString() : null;
        }

        private static string ColLetter(int index)
        {
            string s = "";
            while (index > 0)
            {
                int rem = (index - 1) % 26;
                s = (char)('A' + rem) + s;
                index = (index - 1) / 26;
            }
            return s;
        }

        private Task<SkillResult> CreateComparisonChartAsync(Dictionary<string, object> arguments)
        {
            var categoryColumn = arguments["categoryColumn"].ToString();
            var valueColumns = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["valueColumns"].ToString());
            var chartType = arguments.ContainsKey("chartType") ? arguments["chartType"].ToString().ToLower() : "bar";
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

            var usedRange = sheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;

            var columnMap = new Dictionary<string, int>();
            for (int c = 1; c <= lastCol; c++)
            {
                var colName = sheet.Cells[1, c].Text?.ToString();
                if (!string.IsNullOrEmpty(colName))
                    columnMap[colName] = c;
            }

            if (!columnMap.ContainsKey(categoryColumn))
                return Task.FromResult(new SkillResult { Success = false, Error = $"未找到类别列: {categoryColumn}" });

            var xlChartType = chartType switch
            {
                "line" => Excel.XlChartType.xlLine,
                "radar" => Excel.XlChartType.xlRadar,
                _ => Excel.XlChartType.xlColumnClustered
            };

            var chart = sheet.Shapes.AddChart2(
                Style: -1,
                XlChartType: xlChartType,
                Left: sheet.Cells[1, lastCol + 2].Left,
                Top: sheet.Cells[1, lastCol + 2].Top,
                Width: 500,
                Height: 350
            ).Chart;

            var ranges = new List<string> { $"${categoryColumn}$1:${categoryColumn}${lastRow}" };
            foreach (var col in valueColumns)
            {
                if (columnMap.ContainsKey(col))
                    ranges.Add($"${col}$1:${col}${lastRow}");
            }

            var dataRange = sheet.Range[string.Join(",", ranges)];
            chart.SetSourceData(dataRange);

            chart.HasTitle = true;
            chart.ChartTitle.Text = "对比图表";

            return Task.FromResult(new SkillResult { Success = true, Content = $"对比图表已创建\n类别: {categoryColumn}\n数值列: {string.Join(", ", valueColumns)}" });
        }

        private Task<SkillResult> CreateParetoChartAsync(Dictionary<string, object> arguments)
        {
            var categoryColumn = arguments["categoryColumn"].ToString();
            var valueColumn = arguments["valueColumn"].ToString();
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

            var usedRange = sheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;

            var columnMap = new Dictionary<string, int>();
            for (int c = 1; c <= lastCol; c++)
            {
                var colName = sheet.Cells[1, c].Text?.ToString();
                if (!string.IsNullOrEmpty(colName))
                    columnMap[colName] = c;
            }

            if (!columnMap.ContainsKey(categoryColumn) || !columnMap.ContainsKey(valueColumn))
                return Task.FromResult(new SkillResult { Success = false, Error = "未找到指定的列" });

            int catColIdx = columnMap[categoryColumn];
            int valColIdx = columnMap[valueColumn];

            var data = new List<(string Category, double Value)>();
            for (int r = 2; r <= lastRow; r++)
            {
                var cat = sheet.Cells[r, catColIdx].Text?.ToString();
                var val = GetNumericValue(sheet.Cells[r, valColIdx]);
                if (!string.IsNullOrEmpty(cat))
                    data.Add((cat, val));
            }

            data = data.OrderByDescending(x => x.Value).ToList();

            var existingNames = new List<string>();
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            string paretoSheetName = "帕累托_" + DateTime.Now.ToString("HHmmss");
            int paretoSuffix = 2;
            while (existingNames.Any(n => string.Equals(n, paretoSheetName, StringComparison.OrdinalIgnoreCase)))
            {
                paretoSheetName = $"帕累托_{DateTime.Now.ToString("HHmmss")}_{paretoSuffix}";
                paretoSuffix++;
            }

            var paretoSheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            paretoSheet.Name = paretoSheetName;

            paretoSheet.Cells[1, 1].Value = categoryColumn;
            paretoSheet.Cells[1, 2].Value = valueColumn;
            paretoSheet.Cells[1, 3].Value = "累计占比";

            double total = data.Sum(x => x.Value);
            double cumulative = 0;

            for (int i = 0; i < data.Count; i++)
            {
                paretoSheet.Cells[i + 2, 1].Value = data[i].Category;
                paretoSheet.Cells[i + 2, 2].Value = data[i].Value;
                cumulative += data[i].Value;
                paretoSheet.Cells[i + 2, 3].Value = cumulative / total;
            }

            paretoSheet.UsedRange.Columns.AutoFit();

            var chart = paretoSheet.Shapes.AddChart2(
                Style: -1,
                XlChartType: Excel.XlChartType.xlColumnClustered,
                Left: paretoSheet.Range["E1"].Left,
                Top: paretoSheet.Range["E1"].Top,
                Width: 500,
                Height: 350
            ).Chart;

            chart.SetSourceData(paretoSheet.Range[$"A1:C{data.Count + 1}"]);
            chart.HasTitle = true;
            chart.ChartTitle.Text = "帕累托图";

            paretoSheet.Activate();

            return Task.FromResult(new SkillResult { Success = true, Content = $"帕累托图已创建\n数据已写入工作表: {paretoSheet.Name}" });
        }

        private Task<SkillResult> CreateHistogramAsync(Dictionary<string, object> arguments)
        {
            var valueColumn = arguments["valueColumn"].ToString();
            var binCount = arguments.ContainsKey("binCount") ? Convert.ToInt32(arguments["binCount"]) : 10;
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

            var usedRange = sheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;

            var columnMap = new Dictionary<string, int>();
            for (int c = 1; c <= lastCol; c++)
            {
                var colName = sheet.Cells[1, c].Text?.ToString();
                if (!string.IsNullOrEmpty(colName))
                    columnMap[colName] = c;
            }

            if (!columnMap.ContainsKey(valueColumn))
                return Task.FromResult(new SkillResult { Success = false, Error = $"未找到数值列: {valueColumn}" });

            int valColIdx = columnMap[valueColumn];

            var values = new List<double>();
            for (int r = 2; r <= lastRow; r++)
            {
                var val = GetNumericValue(sheet.Cells[r, valColIdx]);
                if (!double.IsNaN(val))
                    values.Add(val);
            }

            if (values.Count == 0)
                return Task.FromResult(new SkillResult { Success = false, Error = "没有有效的数值数据" });

            var min = values.Min();
            var max = values.Max();
            var binWidth = (max - min) / binCount;

            var existingNames = new List<string>();
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            string histSheetName = "直方图_" + DateTime.Now.ToString("HHmmss");
            int histSuffix = 2;
            while (existingNames.Any(n => string.Equals(n, histSheetName, StringComparison.OrdinalIgnoreCase)))
            {
                histSheetName = $"直方图_{DateTime.Now.ToString("HHmmss")}_{histSuffix}";
                histSuffix++;
            }

            var histSheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            histSheet.Name = histSheetName;

            histSheet.Cells[1, 1].Value = "区间";
            histSheet.Cells[1, 2].Value = "频次";

            var bins = new int[binCount];
            foreach (var val in values)
            {
                int binIndex = Math.Min((int)((val - min) / binWidth), binCount - 1);
                bins[binIndex]++;
            }

            for (int i = 0; i < binCount; i++)
            {
                var lower = min + i * binWidth;
                var upper = min + (i + 1) * binWidth;
                histSheet.Cells[i + 2, 1].Value = $"{lower:F2}-{upper:F2}";
                histSheet.Cells[i + 2, 2].Value = bins[i];
            }

            histSheet.UsedRange.Columns.AutoFit();

            var chart = histSheet.Shapes.AddChart2(
                Style: -1,
                XlChartType: Excel.XlChartType.xlColumnClustered,
                Left: histSheet.Range["D1"].Left,
                Top: histSheet.Range["D1"].Top,
                Width: 400,
                Height: 300
            ).Chart;

            chart.SetSourceData(histSheet.Range[$"A1:B{binCount + 1}"]);
            chart.HasTitle = true;
            chart.ChartTitle.Text = $"{valueColumn} 直方图";

            histSheet.Activate();

            return Task.FromResult(new SkillResult { Success = true, Content = $"直方图已创建\n分组数: {binCount}\n数据已写入工作表: {histSheet.Name}" });
        }

        private Task<SkillResult> CreateBoxPlotAsync(Dictionary<string, object> arguments)
        {
            var valueColumns = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["valueColumns"].ToString());
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

            var usedRange = sheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;

            var columnMap = new Dictionary<string, int>();
            for (int c = 1; c <= lastCol; c++)
            {
                var colName = sheet.Cells[1, c].Text?.ToString();
                if (!string.IsNullOrEmpty(colName))
                    columnMap[colName] = c;
            }

            var existingNames = new List<string>();
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            string boxSheetName = "箱线图_" + DateTime.Now.ToString("HHmmss");
            int boxSuffix = 2;
            while (existingNames.Any(n => string.Equals(n, boxSheetName, StringComparison.OrdinalIgnoreCase)))
            {
                boxSheetName = $"箱线图_{DateTime.Now.ToString("HHmmss")}_{boxSuffix}";
                boxSuffix++;
            }

            var boxSheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            boxSheet.Name = boxSheetName;

            boxSheet.Cells[1, 1].Value = "列名";
            boxSheet.Cells[1, 2].Value = "最小值";
            boxSheet.Cells[1, 3].Value = "Q1";
            boxSheet.Cells[1, 4].Value = "中位数";
            boxSheet.Cells[1, 5].Value = "Q3";
            boxSheet.Cells[1, 6].Value = "最大值";

            int row = 2;
            foreach (var col in valueColumns)
            {
                if (!columnMap.ContainsKey(col)) continue;

                var values = new List<double>();
                for (int r = 2; r <= lastRow; r++)
                {
                    var val = GetNumericValue(sheet.Cells[r, columnMap[col]]);
                    if (!double.IsNaN(val))
                        values.Add(val);
                }

                if (values.Count == 0) continue;

                values.Sort();

                var min = values[0];
                var max = values[values.Count - 1];
                var median = GetPercentile(values, 0.5);
                var q1 = GetPercentile(values, 0.25);
                var q3 = GetPercentile(values, 0.75);

                boxSheet.Cells[row, 1].Value = col;
                boxSheet.Cells[row, 2].Value = min;
                boxSheet.Cells[row, 3].Value = q1;
                boxSheet.Cells[row, 4].Value = median;
                boxSheet.Cells[row, 5].Value = q3;
                boxSheet.Cells[row, 6].Value = max;
                row++;
            }

            boxSheet.UsedRange.Columns.AutoFit();
            boxSheet.Activate();

            return Task.FromResult(new SkillResult { Success = true, Content = $"箱线图数据已生成\n数据已写入工作表: {boxSheet.Name}" });
        }

        private double GetNumericValue(Excel.Range cell)
        {
            try
            {
                var value = cell.Value2;
                if (value == null) return double.NaN;
                if (value is double d) return d;
                if (double.TryParse(value.ToString(), out double result)) return result;
                return double.NaN;
            }
            catch
            {
                return double.NaN;
            }
        }

        private double GetPercentile(List<double> sortedValues, double percentile)
        {
            if (sortedValues.Count == 0) return 0;
            if (sortedValues.Count == 1) return sortedValues[0];

            double index = percentile * (sortedValues.Count - 1);
            int lower = (int)Math.Floor(index);
            int upper = (int)Math.Ceiling(index);

            if (lower == upper)
                return sortedValues[lower];

            return sortedValues[lower] + (sortedValues[upper] - sortedValues[lower]) * (index - lower);
        }
    }
}
