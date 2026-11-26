using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelDatabaseSkill : ISkill
    {
        public string Name => "ExcelDatabase";
        public string Description => "数据库连接技能，支持连接多种数据库并提取数据到Excel";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "connect_database",
                    Description = "连接数据库并获取表名列表。支持SQL Server、MySQL、PostgreSQL、Access、SQLite等数据库。当用户要求连接数据库时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "dbType", new { type = "string", description = "数据库类型：sqlserver/mysql/postgresql/access/sqlite" } },
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
                new SkillTool
                {
                    Name = "execute_query",
                    Description = "执行SQL查询并将结果写入Excel。当用户要求查询数据库、执行SQL时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "dbType", new { type = "string", description = "数据库类型：sqlserver/mysql/postgresql/access/sqlite" } },
                                { "connectionString", new { type = "string", description = "数据库连接字符串" } },
                                { "query", new { type = "string", description = "SQL查询语句" } },
                                { "outputSheetName", new { type = "string", description = "输出工作表名称（可选，默认'查询结果'）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "dbType", "connectionString", "query" }
                },
                new SkillTool
                {
                    Name = "export_table_to_excel",
                    Description = "将数据库表导出到Excel工作表。当用户要求导出数据库表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "dbType", new { type = "string", description = "数据库类型" } },
                                { "connectionString", new { type = "string", description = "数据库连接字符串" } },
                                { "tableName", new { type = "string", description = "要导出的表名" } },
                                { "outputSheetName", new { type = "string", description = "输出工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "dbType", "connectionString", "tableName" }
                },
                new SkillTool
                {
                    Name = "get_table_structure",
                    Description = "获取数据库表的结构信息（字段名、类型等）。当用户要求查看表结构时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "dbType", new { type = "string", description = "数据库类型" } },
                                { "connectionString", new { type = "string", description = "数据库连接字符串" } },
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
                switch (toolName)
                {
                    case "connect_database":
                        return await ConnectDatabaseAsync(arguments);
                    case "execute_query":
                        return await ExecuteQueryAsync(arguments);
                    case "export_table_to_excel":
                        return await ExportTableToExcelAsync(arguments);
                    case "get_table_structure":
                        return await GetTableStructureAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelDatabaseSkill" };
                }
            }
            catch (Exception ex)
            {
                LogError(toolName, ex);
                return new SkillResult { Success = false, Error = "数据库操作失败：" + ex.Message };
            }
        }

        private async Task<SkillResult> ConnectDatabaseAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var dbType = arguments["dbType"].ToString().ToLower();
                var connectionString = arguments["connectionString"].ToString();
                if (dbType == "postgresql")
                    connectionString = BuildPostgreSqlConnectionString(arguments, connectionString);

                try
                {
                    var tableNames = GetTableNames(dbType, connectionString);
                    return new SkillResult 
                    { 
                        Success = true, 
                        Content = $"数据库连接成功，共 {tableNames.Count} 张表：\n{string.Join("\n", tableNames)}" 
                    };
                }
                catch (Exception ex)
                {
                    return new SkillResult { Success = false, Error = $"数据库连接失败：{ex.Message}" };
                }
            });
        }

        private async Task<SkillResult> ExecuteQueryAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var dbType = arguments["dbType"].ToString().ToLower();
                var connectionString = arguments["connectionString"].ToString();
                if (dbType == "postgresql")
                    connectionString = BuildPostgreSqlConnectionString(arguments, connectionString);
                var query = arguments["query"].ToString();
                var outputSheetName = arguments.ContainsKey("outputSheetName") 
                    ? arguments["outputSheetName"].ToString() 
                    : "查询结果";

                try
                {
                    var dataTable = ExecuteQuery(dbType, connectionString, query);
                    WriteDataTableToExcel(dataTable, outputSheetName);
                    return new SkillResult 
                    { 
                        Success = true, 
                        Content = $"查询执行成功，共 {dataTable.Rows.Count} 行数据已写入工作表 '{outputSheetName}'" 
                    };
                }
                catch (Exception ex)
                {
                    return new SkillResult { Success = false, Error = $"查询执行失败：{ex.Message}" };
                }
            });
        }

        private async Task<SkillResult> ExportTableToExcelAsync(Dictionary<string, object> arguments)
        {
            var dbType = arguments["dbType"].ToString().ToLower();
            var connectionString = arguments["connectionString"].ToString();
            if (dbType == "postgresql")
                connectionString = BuildPostgreSqlConnectionString(arguments, connectionString);
            var tableName = arguments["tableName"].ToString();
            var outputSheetName = arguments.ContainsKey("outputSheetName") 
                ? arguments["outputSheetName"].ToString() 
                : tableName;

            return await ExecuteQueryAsync(new Dictionary<string, object>
            {
                { "dbType", dbType },
                { "connectionString", connectionString },
                { "query", $"SELECT * FROM {QuoteTableName(dbType, tableName)}" },
                { "outputSheetName", outputSheetName }
            });
        }

        private async Task<SkillResult> GetTableStructureAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var dbType = arguments["dbType"].ToString().ToLower();
                var connectionString = arguments["connectionString"].ToString();
                if (dbType == "postgresql")
                    connectionString = BuildPostgreSqlConnectionString(arguments, connectionString);
                var tableName = arguments["tableName"].ToString();
                string pgTableName = tableName.Contains('.') ? tableName.Split('.').Last() : tableName;

                try
                {
                    var query = dbType switch
                    {
                        "sqlserver" => $"SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tableName.Replace("'", "''")}'",
                        "mysql" => $"DESCRIBE {QuoteTableName(dbType, tableName)}",
                        "postgresql" => $"SELECT column_name, data_type, character_maximum_length, is_nullable FROM information_schema.columns WHERE table_name = '{pgTableName.Replace("'", "''")}' ORDER BY ordinal_position",
                        "access" or "sqlite" => $"PRAGMA table_info({tableName})",
                        _ => throw new NotSupportedException($"不支持的数据库类型: {dbType}")
                    };

                    var dataTable = ExecuteQuery(dbType, connectionString, query);
                    WriteDataTableToExcel(dataTable, $"{tableName}_结构");
                    return new SkillResult 
                    { 
                        Success = true, 
                        Content = $"表 '{tableName}' 结构已导出，共 {dataTable.Rows.Count} 个字段" 
                    };
                }
                catch (Exception ex)
                {
                    return new SkillResult { Success = false, Error = $"获取表结构失败：{ex.Message}" };
                }
            });
        }

        private List<string> GetTableNames(string dbType, string connectionString)
        {
            var tableNames = new List<string>();

            switch (dbType)
            {
                case "access":
                    using (var conn = new OleDbConnection(connectionString))
                    {
                        conn.Open();
                        var schema = conn.GetSchema("Tables");
                        foreach (DataRow row in schema.Rows)
                        {
                            if (row["TABLE_TYPE"].ToString() == "TABLE")
                                tableNames.Add(row["TABLE_NAME"].ToString());
                        }
                    }
                    break;

                case "sqlite":
                    using (var conn = new SQLiteConnection(connectionString))
                    {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table'", conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                                tableNames.Add(reader["name"].ToString());
                        }
                    }
                    break;

                case "sqlserver":
                case "mysql":
                    var dt = ExecuteQuery(dbType, connectionString, "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'");
                    foreach (DataRow row in dt.Rows)
                        tableNames.Add(row[0].ToString());
                    break;

                case "postgresql":
                    // 排除系统 schema，非 public schema 用 schema.table 形式显示（与 VSTO 版一致）
                    const string pgSql = @"
                        SELECT table_schema, table_name
                        FROM information_schema.tables
                        WHERE table_type IN ('BASE TABLE', 'VIEW')
                          AND table_schema NOT IN ('pg_catalog', 'information_schema')
                        ORDER BY table_schema, table_name";
                    var pgDt = ExecuteQuery(dbType, connectionString, pgSql);
                    foreach (DataRow row in pgDt.Rows)
                    {
                        string schema = row[0].ToString() ?? "";
                        string name = row[1].ToString() ?? "";
                        tableNames.Add(schema == "public" ? name : schema + "." + name);
                    }
                    break;

                default:
                    throw new NotSupportedException($"不支持的数据库类型: {dbType}");
            }

            return tableNames;
        }

        private DataTable ExecuteQuery(string dbType, string connectionString, string query)
        {
            var dataTable = new DataTable();

            switch (dbType)
            {
                case "access":
                    using (var conn = new OleDbConnection(connectionString))
                    using (var adapter = new OleDbDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                case "sqlite":
                    using (var conn = new SQLiteConnection(connectionString))
                    using (var adapter = new SQLiteDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                case "sqlserver":
                    using (var conn = new System.Data.SqlClient.SqlConnection(connectionString))
                    using (var adapter = new System.Data.SqlClient.SqlDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                case "mysql":
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection(connectionString))
                    using (var adapter = new MySql.Data.MySqlClient.MySqlDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                case "postgresql":
                    using (var conn = new Npgsql.NpgsqlConnection(connectionString))
                    using (var adapter = new Npgsql.NpgsqlDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                default:
                    throw new NotSupportedException($"不支持的数据库类型: {dbType}");
            }

            return dataTable;
        }

        private void WriteDataTableToExcel(DataTable dataTable, string sheetName)
        {
            var workbook = ThisAddIn.app.ActiveWorkbook;

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            Excel.Worksheet sheet;
            
            var existingNames = new List<string>();
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            
            string actualSheetName = sheetName;
            if (existingNames.Any(n => string.Equals(n, sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                int suffix = 2;
                while (existingNames.Any(n => string.Equals(n, $"{sheetName}_{suffix}", StringComparison.OrdinalIgnoreCase)))
                {
                    suffix++;
                }
                actualSheetName = $"{sheetName}_{suffix}";
            }
            
            sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            sheet.Name = actualSheetName;

            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                sheet.Cells[1, c + 1].Value = dataTable.Columns[c].ColumnName;
                sheet.Cells[1, c + 1].Font.Bold = true;
            }

            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    sheet.Cells[r + 2, c + 1].Value = dataTable.Rows[r][c];
                }
            }

            sheet.UsedRange.Columns.AutoFit();
            sheet.Activate();

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;
        }

        private static readonly object _logLock = new object();

        /// <summary>记录数据库错误到 %LOCALAPPDATA%\TableMagic\logs\database.log（与 VSTO 主程序一致）</summary>
        private static void LogError(string context, Exception ex)
        {
            try
            {
                string logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TableMagic", "logs");
                Directory.CreateDirectory(logDir);
                string content = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Skill {context}{Environment.NewLine}{ex}{Environment.NewLine}{Environment.NewLine}";
                lock (_logLock)
                {
                    File.AppendAllText(Path.Combine(logDir, "database.log"), content, Encoding.UTF8);
                }
            }
            catch { /* 日志失败不能影响主流程 */ }
        }

        /// <summary>根据 dbType 生成带正确引号的表名（PostgreSQL 用双引号，SQL Server/MySQL 用方括号）</summary>
        private static string QuoteTableName(string dbType, string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName)) return tableName;
            if (dbType == "postgresql")
            {
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
            return tableName;
        }

        /// <summary>由分散字段构建 PostgreSQL 连接串（与 VSTO 版 BuildPostgreSqlConnectionString 对齐）</summary>
        private static string BuildPostgreSqlConnectionString(Dictionary<string, object> args, string fallback)
        {
            string host = args.ContainsKey("host") ? args["host"]?.ToString() : null;
            string database = args.ContainsKey("database") ? args["database"]?.ToString() : null;
            string user = args.ContainsKey("user") ? args["user"]?.ToString() : null;
            string password = args.ContainsKey("password") ? args["password"]?.ToString() : null;
            if (string.IsNullOrWhiteSpace(host) && string.IsNullOrWhiteSpace(database))
                return fallback;

            var csb = new Npgsql.NpgsqlConnectionStringBuilder
            {
                Host = host ?? string.Empty,
                Username = user ?? string.Empty,
                Password = password ?? string.Empty,
                Database = database ?? string.Empty,
                Timeout = 15,
                CommandTimeout = 60
            };
            int port;
            string portStr = args.ContainsKey("port") ? args["port"]?.ToString() : null;
            if (int.TryParse(portStr, out port) && port > 0 && port <= 65535)
                csb.Port = port;
            return csb.ConnectionString;
        }
    }
}
