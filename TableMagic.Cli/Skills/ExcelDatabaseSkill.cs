using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelDatabaseSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelDatabaseSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelDatabase";
    public string Description => "数据库连接技能，支持SQLite/MySQL/PostgreSQL/SQL Server";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "connect_database",
                Description = "连接数据库并获取表名列表。支持sqlserver/mysql/postgresql/sqlite",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "dbType", new { type = "string", description = "数据库类型：sqlserver/mysql/postgresql/sqlite" } },
                            { "connectionString", new { type = "string", description = "数据库连接字符串或文件路径（PostgreSQL 也可用下面的分散字段）" } },
                            { "host", new { type = "string", description = "PostgreSQL 主机名（可选，提供时优先于 connectionString）" } },
                            { "port", new { type = "string", description = "PostgreSQL 端口，默认5432（可选）" } },
                            { "database", new { type = "string", description = "PostgreSQL 数据库名（可选）" } },
                            { "user", new { type = "string", description = "PostgreSQL 用户名（可选）" } },
                            { "password", new { type = "string", description = "PostgreSQL 密码（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "dbType", "connectionString" }
            },
            new()
            {
                Name = "execute_query",
                Description = "执行SQL查询并将结果写入Excel",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "dbType", new { type = "string", description = "数据库类型" } },
                            { "connectionString", new { type = "string", description = "连接字符串" } },
                            { "query", new { type = "string", description = "SQL查询语句" } },
                            { "outputFileName", new { type = "string", description = "输出工作簿文件名（可选）" } },
                            { "outputSheetName", new { type = "string", description = "输出工作表名称（默认'查询结果'）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "dbType", "connectionString", "query" }
            },
            new()
            {
                Name = "export_table_to_excel",
                Description = "将数据库表导出到Excel工作表",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "dbType", new { type = "string", description = "数据库类型" } },
                            { "connectionString", new { type = "string", description = "连接字符串" } },
                            { "tableName", new { type = "string", description = "要导出的表名" } },
                            { "outputFileName", new { type = "string", description = "输出工作簿文件名（可选）" } },
                            { "outputSheetName", new { type = "string", description = "输出工作表名称（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "dbType", "connectionString", "tableName" }
            },
            new()
            {
                Name = "get_table_structure",
                Description = "获取数据库表的结构信息（字段名、类型等）",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "dbType", new { type = "string", description = "数据库类型" } },
                            { "connectionString", new { type = "string", description = "连接字符串" } },
                            { "tableName", new { type = "string", description = "表名" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "dbType", "connectionString", "tableName" }
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "connect_database" => await ConnectDatabaseAsync(arguments),
                "execute_query" => await ExecuteQueryAsync(arguments),
                "export_table_to_excel" => await ExportTableAsync(arguments),
                "get_table_structure" => await GetTableStructureAsync(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex)
        {
            LogError(toolName, ex);
            return new SkillResult { Success = false, Error = "数据库操作失败：" + ex.Message };
        }
    }

    private async Task<SkillResult> ConnectDatabaseAsync(Dictionary<string, object> args)
    {
        return await Task.Run(() =>
        {
            var dbType = GetStr(args, "dbType")?.ToLower();
            var connStr = GetStr(args, "connectionString");
            if (dbType == "postgresql")
                connStr = BuildPostgreSqlConnectionString(args, connStr);
            var tables = GetTableNames(dbType!, connStr!, args);
            return SkillResult.Ok($"数据库连接成功，共 {tables.Count} 张表：\n{string.Join("\n", tables)}");
        });
    }

    private async Task<SkillResult> ExecuteQueryAsync(Dictionary<string, object> args)
    {
        return await Task.Run(() =>
        {
            var dbType = GetStr(args, "dbType")?.ToLower();
            var connStr = GetStr(args, "connectionString");
            if (dbType == "postgresql")
                connStr = BuildPostgreSqlConnectionString(args, connStr);
            var query = GetStr(args, "query");
            var outputFn = GetStr(args, "outputFileName");
            var outputSn = GetStr(args, "outputSheetName") ?? "查询结果";

            var dt = ExecuteDbQuery(dbType!, connStr!, query!);
            WriteDataTableToExcel(dt, outputFn, outputSn);
            return SkillResult.Ok($"查询执行成功，共 {dt.Rows.Count} 行数据已写入工作表 '{outputSn}'");
        });
    }

    private async Task<SkillResult> ExportTableAsync(Dictionary<string, object> args)
    {
        var dbType = GetStr(args, "dbType")?.ToLower();
        var tableName = GetStr(args, "tableName");
        var newArgs = new Dictionary<string, object>(args)
        {
            ["query"] = $"SELECT * FROM {QuoteTableName(dbType!, tableName!)}"
        };
        if (!newArgs.ContainsKey("outputSheetName"))
            newArgs["outputSheetName"] = GetStr(args, "tableName") ?? "导出表";
        return await ExecuteQueryAsync(newArgs);
    }

    private async Task<SkillResult> GetTableStructureAsync(Dictionary<string, object> args)
    {
        return await Task.Run(() =>
        {
            var dbType = GetStr(args, "dbType")?.ToLower();
            var connStr = GetStr(args, "connectionString");
            if (dbType == "postgresql")
                connStr = BuildPostgreSqlConnectionString(args, connStr);
            var tableName = GetStr(args, "tableName");
            string pgTableName = tableName != null && tableName.Contains('.') ? tableName.Split('.').Last() : tableName ?? "";
            var query = dbType switch
            {
                "sqlite" => $"PRAGMA table_info({tableName})",
                "mysql" => $"DESCRIBE {QuoteTableName(dbType, tableName!)}",
                "postgresql" => $"SELECT column_name, data_type, character_maximum_length, is_nullable FROM information_schema.columns WHERE table_name = '{pgTableName?.Replace("'", "''")}' ORDER BY ordinal_position",
                "sqlserver" => $"SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tableName?.Replace("'", "''")}'",
                _ => throw new NotSupportedException($"不支持的数据库类型: {dbType}")
            };
            var dt = ExecuteDbQuery(dbType, connStr, query);
            var outputFn = GetStr(args, "outputFileName");
            WriteDataTableToExcel(dt, outputFn, $"{tableName}_结构");
            return SkillResult.Ok($"表 '{tableName}' 结构已导出，共 {dt.Rows.Count} 个字段");
        });
    }

    private List<string> GetTableNames(string dbType, string connStr, Dictionary<string, object>? args = null)
    {
        var tables = new List<string>();
        switch (dbType)
        {
            case "sqlite":
                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using var cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table'", conn);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read()) tables.Add(reader["name"].ToString()!);
                }
                break;
            case "postgresql":
                // 排除系统 schema，非 public schema 用 schema.table 形式显示（与 VSTO 版一致）
                const string pgSql = @"
                    SELECT table_schema, table_name
                    FROM information_schema.tables
                    WHERE table_type IN ('BASE TABLE', 'VIEW')
                      AND table_schema NOT IN ('pg_catalog', 'information_schema')
                    ORDER BY table_schema, table_name";
                var pgDt = ExecuteDbQuery(dbType, connStr, pgSql);
                foreach (DataRow row in pgDt.Rows)
                {
                    string schema = row[0].ToString() ?? "";
                    string name = row[1].ToString() ?? "";
                    tables.Add(schema == "public" ? name : schema + "." + name);
                }
                break;
            case "mysql":
            case "sqlserver":
                var dt = ExecuteDbQuery(dbType, connStr, "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'");
                foreach (DataRow row in dt.Rows) tables.Add(row[0].ToString()!);
                break;
            default: throw new NotSupportedException($"不支持的数据库类型: {dbType}");
        }
        return tables;
    }

    private DataTable ExecuteDbQuery(string dbType, string connStr, string query)
    {
        var dt = new DataTable();
        switch (dbType)
        {
            case "sqlite":
                using (var conn = new SQLiteConnection(connStr))
                using (var adapter = new SQLiteDataAdapter(query, conn)) { adapter.Fill(dt); }
                break;
            case "mysql":
                using (var conn = new MySql.Data.MySqlClient.MySqlConnection(connStr))
                using (var adapter = new MySql.Data.MySqlClient.MySqlDataAdapter(query, conn)) { adapter.Fill(dt); }
                break;
            case "postgresql":
                using (var conn = new Npgsql.NpgsqlConnection(connStr))
                {
                    conn.Open();
                    using var adapter = new Npgsql.NpgsqlDataAdapter(query, conn);
                    adapter.Fill(dt);
                }
                break;
            case "sqlserver":
                using (var conn = new Microsoft.Data.SqlClient.SqlConnection(connStr))
                using (var adapter = new Microsoft.Data.SqlClient.SqlDataAdapter(query, conn)) { adapter.Fill(dt); }
                break;
            default: throw new NotSupportedException($"不支持的数据库类型: {dbType}");
        }
        return dt;
    }

    private void WriteDataTableToExcel(DataTable dt, string fileName, string sheetName)
    {
        var fn = fileName ?? _provider.GetOpenWorkbooks().FirstOrDefault();
        if (fn == null)
        {
            fn = "query_result.xlsx";
            _provider.CreateWorkbook(fn, sheetName);
        }
        else
        {
            if (!_provider.GetOpenWorkbooks().Contains(fn))
                _provider.OpenWorkbook(fn);
            _provider.CreateWorksheet(fn, sheetName);
        }

        for (int c = 0; c < dt.Columns.Count; c++)
            _provider.SetCellValue(fn, sheetName, 1, c + 1, dt.Columns[c].ColumnName);

        for (int r = 0; r < dt.Rows.Count; r++)
            for (int c = 0; c < dt.Columns.Count; c++)
                _provider.SetCellValue(fn, sheetName, r + 2, c + 1, dt.Rows[r][c]);

        _provider.SaveWorkbook(fn);
    }

    /// <summary>记录数据库错误到 %LOCALAPPDATA%\TableMagic\logs\database.log（不打断主流程）</summary>
    private static void LogError(string context, Exception ex)
    {
        try
        {
            string logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TableMagic", "logs");
            Directory.CreateDirectory(logDir);
            string content = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] CLI {context}{Environment.NewLine}{ex}{Environment.NewLine}{Environment.NewLine}";
            lock (_logLock)
            {
                File.AppendAllText(Path.Combine(logDir, "database.log"), content, Encoding.UTF8);
            }
        }
        catch { /* 日志失败不能影响主流程 */ }
    }

    private static readonly object _logLock = new object();

    /// <summary>根据 dbType 生成带正确引号的表名（PostgreSQL 用双引号，SQL Server 用方括号）</summary>
    private static string QuoteTableName(string dbType, string tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) return tableName;
        if (dbType == "postgresql")
        {
            // 非 public schema 名如 sales.orders，需逐段加引号；自动兼容已带引号的情况
            var parts = tableName.Trim().Trim('"').Split('.');
            for (int i = 0; i < parts.Length; i++)
                parts[i] = "\"" + parts[i].Replace("\"", "\"\"") + "\"";
            return string.Join(".", parts);
        }
        if (dbType == "sqlserver" || dbType == "mysql")
        {
            var parts = tableName.Trim().Trim('[', ']').Split('.');
            for (int i = 0; i < parts.Length; i++)
                parts[i] = "[" + parts[i].Replace("]", "]]") + "]";
            return string.Join(".", parts);
        }
        return tableName; // SQLite 不加引号时会原样传递
    }

    /// <summary>由分散字段构建 PostgreSQL 连接串（与 VSTO 版 BuildPostgreSqlConnectionString 对齐）</summary>
    private static string BuildPostgreSqlConnectionString(Dictionary<string, object> args, string fallback)
    {
        string host = GetStr(args, "host");
        string database = GetStr(args, "database");
        string user = GetStr(args, "user");
        string password = GetStr(args, "password");
        int port;
        if (string.IsNullOrWhiteSpace(host) && string.IsNullOrWhiteSpace(database))
            return fallback; // 没有分散字段时使用传入的原始连接串

        var csb = new Npgsql.NpgsqlConnectionStringBuilder
        {
            Host = host ?? string.Empty,
            Username = user ?? string.Empty,
            Password = password ?? string.Empty,
            Database = database ?? string.Empty,
            Timeout = 15,
            CommandTimeout = 60
        };
        if (int.TryParse(GetStr(args, "port"), out port) && port > 0 && port <= 65535)
            csb.Port = port;
        return csb.ConnectionString;
    }

    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
}