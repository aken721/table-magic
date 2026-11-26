using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using SkiaSharp;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelChartEnhanceSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelChartEnhanceSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelChartEnhance";
    public string Description => "增强图表：词云、动态图、帕累托、直方图、箱线图、对比图";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "create_word_cloud", Description = "根据文本数据生成词云图片",
                Parameters = P(new[]{"textColumn"}, new[]{"fileName","sheetName","maxWords","width","height","outputPath"}), RequiredParameters = new List<string>{"textColumn"} },
            new() { Name = "create_dynamic_chart", Description = "创建动态图表。支持二维表(categoryColumn/valueColumn)和一维表(按分组列自动透视成二维表再生成图表)",
                Parameters = P(new[]{"categoryColumn","valueColumn"}, new[]{"fileName","sheetName","chartType","title","tableType","dataOrientation","skipHeader","groupColumn","itemColumn","valueColumn1d"}), RequiredParameters = new List<string>{"categoryColumn","valueColumn"} },
            new() { Name = "create_comparison_chart", Description = "创建对比图（多系列数据对比）",
                Parameters = P(new[]{"categoryColumn","valueColumns"}, new[]{"fileName","sheetName","chartType","title"}), RequiredParameters = new List<string>{"categoryColumn","valueColumns"} },
            new() { Name = "create_pareto_chart", Description = "创建帕累托图（二八分析）",
                Parameters = P(new[]{"categoryColumn","valueColumn"}, new[]{"fileName","sheetName"}), RequiredParameters = new List<string>{"categoryColumn","valueColumn"} },
            new() { Name = "create_histogram", Description = "创建直方图",
                Parameters = P(new[]{"valueColumn"}, new[]{"fileName","sheetName","binCount"}), RequiredParameters = new List<string>{"valueColumn"} },
            new() { Name = "create_box_plot", Description = "创建箱线图（五数概括）",
                Parameters = P(new[]{"valueColumns"}, new[]{"fileName","sheetName"}), RequiredParameters = new List<string>{"valueColumns"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "create_word_cloud" => CreateWordCloud(arguments),
                "create_dynamic_chart" => CreateDynamicChart(arguments),
                "create_comparison_chart" => CreateComparisonChart(arguments),
                "create_pareto_chart" => CreateParetoChart(arguments),
                "create_histogram" => CreateHistogram(arguments),
                "create_box_plot" => CreateBoxPlot(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult CreateWordCloud(Dictionary<string, object> args)
    {
        var textColumn = GetStr(args, "textColumn");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");
        var maxWords = GetInt(args, "maxWords", 100);
        var width = GetInt(args, "width", 800);
        var height = GetInt(args, "height", 600);
        var outputPath = GetStr(args, "outputPath");

        if (string.IsNullOrEmpty(textColumn))
            return SkillResult.MissingParamsResult("create_word_cloud", new List<MissingParam>
            {
                new() { Name = "textColumn", Description = "文本数据列名", PromptHint = "请提供文本数据列名" }
            });

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);

        var headerRow = BuildHeaderMap(wb, sn, lastCol);
        if (!headerRow.TryGetValue(textColumn, out var textColIdx))
            return SkillResult.FromError($"未找到文本列: {textColumn}");

        var wordWeights = new Dictionary<string, double>();
        for (int r = 2; r <= lastRow; r++)
        {
            var text = _provider.GetCellValue(wb, sn, r, textColIdx)?.ToString();
            if (string.IsNullOrEmpty(text)) continue;
            var words = text.Split(new[] { ' ', ',', '，', '、', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var w in words)
            {
                var trimmed = w.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                if (wordWeights.ContainsKey(trimmed)) wordWeights[trimmed]++;
                else wordWeights[trimmed] = 1;
            }
        }

        var topWords = wordWeights.OrderByDescending(x => x.Value).Take(maxWords).ToList();
        if (topWords.Count == 0)
            return SkillResult.FromError("没有可用的文本数据");

        if (string.IsNullOrEmpty(outputPath))
            outputPath = Path.Combine(Path.GetTempPath(), $"wordcloud_{Guid.NewGuid():N}.png");

        RenderWordCloud(topWords, width, height, outputPath);

        return SkillResult.Ok($"词云已生成，共 {topWords.Count} 个词语\n图片保存到: {outputPath}");
    }

    private void RenderWordCloud(List<KeyValuePair<string, double>> words, int width, int height, string outputPath)
    {
        using var bitmap = new SKBitmap(width, height);
        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.White);

        var random = new Random();
        var maxWeight = words[0].Value;
        var minWeight = words[words.Count - 1].Value;
        var range = maxWeight - minWeight;
        if (range == 0) range = 1;

        var placedRects = new List<SKRect>();

        foreach (var (word, weight) in words)
        {
            var fontSize = (int)(12 + (weight - minWeight) / range * 48);
            using var font = new SKFont(SKTypeface.FromFamilyName("Microsoft YaHei", SKFontStyle.Bold), fontSize);
            using var paint = new SKPaint { Color = GetRandomColor(random), IsAntialias = true };

            var textWidth = font.MeasureText(word);
            var textHeight = fontSize;

            bool placed = false;
            for (int attempt = 0; attempt < 200 && !placed; attempt++)
            {
                var x = random.Next(10, Math.Max(11, width - (int)textWidth - 10));
                var y = random.Next(10, Math.Max(11, height - textHeight - 10));
                var rect = new SKRect(x, y, x + textWidth, y + textHeight);

                bool overlaps = false;
                foreach (var pr in placedRects)
                {
                    if (rect.IntersectsWith(pr)) { overlaps = true; break; }
                }

                if (!overlaps)
                {
#pragma warning disable CS0618
                    canvas.DrawText(word, x, y + textHeight * 0.8f, font, paint);
#pragma warning restore CS0618
                    placedRects.Add(rect);
                    placed = true;
                }
            }
        }

        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        File.WriteAllBytes(outputPath, data.ToArray());
    }

    private static SKColor GetRandomColor(Random random)
    {
        var colors = new[]
        {
            SKColors.DarkBlue, SKColors.DarkRed, SKColors.DarkGreen, SKColors.DarkOrange,
            SKColors.Purple, SKColors.Teal, SKColors.Brown, SKColors.Navy,
            SKColors.Crimson, SKColors.ForestGreen, SKColors.DarkCyan, SKColors.Indigo
        };
        return colors[random.Next(colors.Length)];
    }

    private SkillResult CreateDynamicChart(Dictionary<string, object> args)
    {
        var categoryColumn = GetStr(args, "categoryColumn");
        var valueColumn = GetStr(args, "valueColumn");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");
        var chartType = GetStr(args, "chartType") ?? "column";
        var title = GetStr(args, "title");
        bool oneDimension = string.Equals(GetStr(args, "tableType"), "1d", StringComparison.OrdinalIgnoreCase)
                         || GetStr(args, "tableType") == "1";
        bool skipHeader = GetStr(args, "skipHeader") != "false" && GetStr(args, "skipHeader") != "0";

        if (!oneDimension && (string.IsNullOrEmpty(categoryColumn) || string.IsNullOrEmpty(valueColumn)))
            return SkillResult.FromError("需要提供 categoryColumn 和 valueColumn");

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);

        var headerRow = BuildHeaderMap(wb, sn, lastCol);
        var chartTitle = string.IsNullOrEmpty(title) ? "动态图表" : title;
        string dataRange;
        string usedSheet;

        if (oneDimension)
        {
            // 一维表：按指定的分组列/类别列/数值列读取，透视成二维表后生成图表
            var groupCol = GetStr(args, "groupColumn") ?? categoryColumn;
            var itemCol = GetStr(args, "itemColumn");
            var valCol = GetStr(args, "valueColumn1d") ?? valueColumn;
            if (string.IsNullOrEmpty(groupCol) || string.IsNullOrEmpty(itemCol) || string.IsNullOrEmpty(valCol))
                return SkillResult.FromError("一维表模式需要提供 groupColumn、itemColumn、valueColumn1d");
            if (!headerRow.ContainsKey(groupCol) || !headerRow.ContainsKey(itemCol) || !headerRow.ContainsKey(valCol))
                return SkillResult.FromError("一维表模式下未找到指定的列（groupColumn/itemColumn/valueColumn1d）");

            int gIdx = headerRow[groupCol], iIdx = headerRow[itemCol], vIdx = headerRow[valCol];
            int startRow = skipHeader ? 2 : 1;

            // 读取一维数据并按分组名 → 类别名聚合（同一分组同一类别取最后一条）
            var groups = new Dictionary<string, Dictionary<string, decimal>>();
            var groupOrder = new List<string>();
            var itemOrder = new List<string>();
            for (int r = startRow; r <= lastRow; r++)
            {
                string g = _provider.GetCellValue(wb, sn, r, gIdx)?.ToString()?.Trim() ?? "";
                string it = _provider.GetCellValue(wb, sn, r, iIdx)?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(g) || string.IsNullOrEmpty(it)) continue;
                decimal val;
                string rawVal = _provider.GetCellValue(wb, sn, r, vIdx)?.ToString() ?? "";
                if (!decimal.TryParse(rawVal, out val)) continue;
                if (!groups.ContainsKey(g)) { groups[g] = new Dictionary<string, decimal>(); groupOrder.Add(g); }
                if (!groups[g].ContainsKey(it)) itemOrder.Add(it);
                groups[g][it] = val;
            }
            if (groups.Count == 0)
                return SkillResult.FromError("一维表未读取到有效数据（请检查 skipHeader 与列选择）");

            // 写入临时 sheet 形成二维表：第一行为分组名，第一列为类别名
            string pivotSheet = $"_动态图透视_{DateTime.Now:yyyyMMdd_HHmmss}";
            _provider.CreateWorksheet(fileName!, pivotSheet);
            _provider.SetCellValue(fileName!, pivotSheet, 1, 1, "类别");
            for (int c = 0; c < groupOrder.Count; c++)
                _provider.SetCellValue(fileName!, pivotSheet, 1, c + 2, groupOrder[c]);
            for (int r = 0; r < itemOrder.Count; r++)
            {
                _provider.SetCellValue(fileName!, pivotSheet, r + 2, 1, itemOrder[r]);
                for (int c = 0; c < groupOrder.Count; c++)
                {
                    decimal v;
                    _provider.SetCellValue(fileName!, pivotSheet, r + 2, c + 2,
                        groups[groupOrder[c]].TryGetValue(itemOrder[r], out v) ? (object)v : "");
                }
            }
            usedSheet = pivotSheet;
            dataRange = $"A1:{ColLetter(groupOrder.Count + 1)}{itemOrder.Count + 1}";
        }
        else
        {
            if (!headerRow.ContainsKey(categoryColumn!) || !headerRow.ContainsKey(valueColumn!))
                return SkillResult.FromError("未找到指定的列");
            usedSheet = sn;
            dataRange = $"{categoryColumn}1:{valueColumn}{lastRow}";
        }

        var result = _provider.CreateChart(wb, usedSheet, dataRange, chartType, chartTitle);
        string mode = oneDimension ? "一维表→二维表" : "二维表";
        return SkillResult.Ok($"动态图表已创建（{mode}）\n类型: {chartType}\n标题: {chartTitle}\n数据范围: {dataRange}\n{result}");
    }

    /// <summary>列索引(1-based)转 Excel 列字母</summary>
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

    private SkillResult CreateComparisonChart(Dictionary<string, object> args)
    {
        var categoryColumn = GetStr(args, "categoryColumn");
        var valueColumnsStr = GetStr(args, "valueColumns");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");
        var chartType = GetStr(args, "chartType") ?? "column";
        var title = GetStr(args, "title");

        if (string.IsNullOrEmpty(categoryColumn) || string.IsNullOrEmpty(valueColumnsStr))
            return SkillResult.FromError("需要提供 categoryColumn 和 valueColumns");

        List<string> valueColumns;
        try { valueColumns = JsonSerializer.Deserialize<List<string>>(valueColumnsStr)!; }
        catch { valueColumns = valueColumnsStr.Split(',').Select(s => s.Trim()).ToList(); }

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);

        var headerRow = BuildHeaderMap(wb, sn, lastCol);
        if (!headerRow.ContainsKey(categoryColumn))
            return SkillResult.FromError($"未找到类别列: {categoryColumn}");

        var validCols = valueColumns.Where(c => headerRow.ContainsKey(c)).ToList();
        if (validCols.Count == 0)
            return SkillResult.FromError("未找到任何数值列");

        var dataRange = $"{categoryColumn}1:{validCols[validCols.Count - 1]}{lastRow}";
        var chartTitle = string.IsNullOrEmpty(title) ? "对比图表" : title;
        var result = _provider.CreateChart(wb, sn, dataRange, chartType, chartTitle);

        return SkillResult.Ok($"对比图表已创建\n类别: {categoryColumn}\n数值列: {string.Join(", ", validCols)}\n{result}");
    }

    private SkillResult CreateParetoChart(Dictionary<string, object> args)
    {
        var categoryColumn = GetStr(args, "categoryColumn");
        var valueColumn = GetStr(args, "valueColumn");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");

        if (string.IsNullOrEmpty(categoryColumn) || string.IsNullOrEmpty(valueColumn))
            return SkillResult.FromError("需要提供 categoryColumn 和 valueColumn");

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);

        var headerRow = BuildHeaderMap(wb, sn, lastCol);
        if (!headerRow.TryGetValue(categoryColumn, out var catColIdx) || !headerRow.TryGetValue(valueColumn, out var valColIdx))
            return SkillResult.FromError("未找到指定的列");

        var data = new List<(string Category, double Value)>();
        for (int r = 2; r <= lastRow; r++)
        {
            var cat = _provider.GetCellValue(wb, sn, r, catColIdx)?.ToString();
            var valStr = _provider.GetCellValue(wb, sn, r, valColIdx)?.ToString();
            if (!string.IsNullOrEmpty(cat) && double.TryParse(valStr, out var val))
                data.Add((cat, val));
        }

        data = data.OrderByDescending(x => x.Value).ToList();
        if (data.Count == 0)
            return SkillResult.FromError("没有有效的数据");

        var paretoSheetName = $"帕累托_{DateTime.Now:HHmmss}";
        _provider.CreateWorksheet(wb, paretoSheetName);

        _provider.SetCellValue(wb, paretoSheetName, 1, 1, categoryColumn);
        _provider.SetCellValue(wb, paretoSheetName, 1, 2, valueColumn);
        _provider.SetCellValue(wb, paretoSheetName, 1, 3, "累计占比");

        double total = data.Sum(x => x.Value);
        double cumulative = 0;

        for (int i = 0; i < data.Count; i++)
        {
            _provider.SetCellValue(wb, paretoSheetName, i + 2, 1, data[i].Category);
            _provider.SetCellValue(wb, paretoSheetName, i + 2, 2, data[i].Value);
            cumulative += data[i].Value;
            _provider.SetCellValue(wb, paretoSheetName, i + 2, 3, Math.Round(cumulative / total, 4));
        }

        var dataRange = $"A1:C{data.Count + 1}";
        _provider.CreateChart(wb, paretoSheetName, dataRange, "column", "帕累托图");
        _provider.SaveWorkbook(wb);

        return SkillResult.Ok($"帕累托图已创建\n数据已写入工作表: {paretoSheetName}");
    }

    private SkillResult CreateHistogram(Dictionary<string, object> args)
    {
        var valueColumn = GetStr(args, "valueColumn");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");
        var binCount = GetInt(args, "binCount", 10);

        if (string.IsNullOrEmpty(valueColumn))
            return SkillResult.FromError("需要提供 valueColumn");

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);

        var headerRow = BuildHeaderMap(wb, sn, lastCol);
        if (!headerRow.TryGetValue(valueColumn, out var valColIdx))
            return SkillResult.FromError($"未找到数值列: {valueColumn}");

        var values = new List<double>();
        for (int r = 2; r <= lastRow; r++)
        {
            var valStr = _provider.GetCellValue(wb, sn, r, valColIdx)?.ToString();
            if (double.TryParse(valStr, out var val))
                values.Add(val);
        }

        if (values.Count == 0)
            return SkillResult.FromError("没有有效的数值数据");

        var min = values.Min();
        var max = values.Max();
        var binWidth = (max - min) / binCount;
        if (binWidth == 0) binWidth = 1;

        var histSheetName = $"直方图_{DateTime.Now:HHmmss}";
        _provider.CreateWorksheet(wb, histSheetName);

        _provider.SetCellValue(wb, histSheetName, 1, 1, "区间");
        _provider.SetCellValue(wb, histSheetName, 1, 2, "频次");

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
            _provider.SetCellValue(wb, histSheetName, i + 2, 1, $"{lower:F2}-{upper:F2}");
            _provider.SetCellValue(wb, histSheetName, i + 2, 2, bins[i]);
        }

        var dataRange = $"A1:B{binCount + 1}";
        _provider.CreateChart(wb, histSheetName, dataRange, "column", $"{valueColumn} 直方图");
        _provider.SaveWorkbook(wb);

        return SkillResult.Ok($"直方图已创建\n分组数: {binCount}\n数据已写入工作表: {histSheetName}");
    }

    private SkillResult CreateBoxPlot(Dictionary<string, object> args)
    {
        var valueColumnsStr = GetStr(args, "valueColumns");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");

        if (string.IsNullOrEmpty(valueColumnsStr))
            return SkillResult.FromError("需要提供 valueColumns");

        List<string> valueColumns;
        try { valueColumns = JsonSerializer.Deserialize<List<string>>(valueColumnsStr)!; }
        catch { valueColumns = valueColumnsStr.Split(',').Select(s => s.Trim()).ToList(); }

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);

        var headerRow = BuildHeaderMap(wb, sn, lastCol);

        var boxSheetName = $"箱线图_{DateTime.Now:HHmmss}";
        _provider.CreateWorksheet(wb, boxSheetName);

        _provider.SetCellValue(wb, boxSheetName, 1, 1, "列名");
        _provider.SetCellValue(wb, boxSheetName, 1, 2, "最小值");
        _provider.SetCellValue(wb, boxSheetName, 1, 3, "Q1");
        _provider.SetCellValue(wb, boxSheetName, 1, 4, "中位数");
        _provider.SetCellValue(wb, boxSheetName, 1, 5, "Q3");
        _provider.SetCellValue(wb, boxSheetName, 1, 6, "最大值");

        int row = 2;
        foreach (var col in valueColumns)
        {
            if (!headerRow.TryGetValue(col, out var colIdx)) continue;

            var values = new List<double>();
            for (int r = 2; r <= lastRow; r++)
            {
                var valStr = _provider.GetCellValue(wb, sn, r, colIdx)?.ToString();
                if (double.TryParse(valStr, out var val))
                    values.Add(val);
            }

            if (values.Count == 0) continue;

            values.Sort();
            var min = values[0];
            var max = values[values.Count - 1];
            var median = GetPercentile(values, 0.5);
            var q1 = GetPercentile(values, 0.25);
            var q3 = GetPercentile(values, 0.75);

            _provider.SetCellValue(wb, boxSheetName, row, 1, col);
            _provider.SetCellValue(wb, boxSheetName, row, 2, Math.Round(min, 4));
            _provider.SetCellValue(wb, boxSheetName, row, 3, Math.Round(q1, 4));
            _provider.SetCellValue(wb, boxSheetName, row, 4, Math.Round(median, 4));
            _provider.SetCellValue(wb, boxSheetName, row, 5, Math.Round(q3, 4));
            _provider.SetCellValue(wb, boxSheetName, row, 6, Math.Round(max, 4));
            row++;
        }

        _provider.SaveWorkbook(wb);

        return SkillResult.Ok($"箱线图数据已生成\n数据已写入工作表: {boxSheetName}");
    }

    private string? ResolveWorkbook(string fileName)
    {
        if (!string.IsNullOrEmpty(fileName)) return fileName;
        return _provider.GetOpenWorkbooks().FirstOrDefault();
    }

    private Dictionary<string, int> BuildHeaderMap(string wb, string sn, int lastCol)
    {
        var map = new Dictionary<string, int>();
        for (int c = 1; c <= lastCol; c++)
        {
            var val = _provider.GetCellValue(wb, sn, 1, c)?.ToString();
            if (!string.IsNullOrEmpty(val)) map[val] = c;
        }
        return map;
    }

    private static double GetPercentile(List<double> sortedValues, double percentile)
    {
        if (sortedValues.Count == 0) return 0;
        if (sortedValues.Count == 1) return sortedValues[0];
        double index = percentile * (sortedValues.Count - 1);
        int lower = (int)Math.Floor(index);
        int upper = (int)Math.Ceiling(index);
        if (lower == upper) return sortedValues[lower];
        return sortedValues[lower] + (sortedValues[upper] - sortedValues[lower]) * (index - lower);
    }

    private static Dictionary<string, object> P(string[] req, string[] opt)
    {
        var p = new Dictionary<string, object>();
        foreach (var r in req) p[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in opt) p[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object> { { "type", "object" }, { "properties", p } };
    }
    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
    private static int GetInt(Dictionary<string, object> a, string k, int def = 0) => a.ContainsKey(k) && int.TryParse(a[k]?.ToString(), out var v) ? v : def;
}
