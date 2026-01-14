using System.Globalization;
using System.Text;
using Oracle.ManagedDataAccess.Client;

namespace SqlSdiCli;

internal sealed record ColumnMeta(string Name, string Comment, string DataType);

internal static class Program
{
    private static readonly CultureInfo InvariantCulture = CultureInfo.InvariantCulture;

    private static int Main(string[] args)
    {
        if (args.Length == 0 || args.Contains("--help", StringComparer.OrdinalIgnoreCase))
        {
            PrintUsage();
            return 0;
        }

        var options = ParseArgs(args);
        if (!options.TryGetValue("connection-string", out var connectionString) || string.IsNullOrWhiteSpace(connectionString))
        {
            Console.Error.WriteLine("--connection-string is required.");
            return 1;
        }

        if (!options.TryGetValue("schema", out var schema) || string.IsNullOrWhiteSpace(schema))
        {
            Console.Error.WriteLine("--schema is required.");
            return 1;
        }

        if (!options.TryGetValue("target-column", out var targetColumn) || string.IsNullOrWhiteSpace(targetColumn))
        {
            Console.Error.WriteLine("--target-column is required.");
            return 1;
        }

        if (!options.TryGetValue("condition", out var condition) || string.IsNullOrWhiteSpace(condition))
        {
            Console.Error.WriteLine("--condition is required.");
            return 1;
        }

        var flashback = options.GetValueOrDefault("flashback-timestamp");
        var preSql = options.GetValueOrDefault("pre-sql");
        var outputFile = options.GetValueOrDefault("output-file");
        var splitByOperation = options.ContainsKey("split-by-operation");
        var sqlDirectory = ResolveSqlDirectory(options.GetValueOrDefault("sql-dir"));

        var primaryKeySqlPath = Path.Combine(sqlDirectory, "primary_keys.sql");
        var columnDefinitionsSqlPath = Path.Combine(sqlDirectory, "column_definitions.sql");
        if (!File.Exists(primaryKeySqlPath))
        {
            Console.Error.WriteLine($"Primary key SQL file not found at '{primaryKeySqlPath}'.");
            return 1;
        }

        if (!File.Exists(columnDefinitionsSqlPath))
        {
            Console.Error.WriteLine($"Column definitions SQL file not found at '{columnDefinitionsSqlPath}'.");
            return 1;
        }

        var primaryKeysSql = File.ReadAllText(primaryKeySqlPath);
        var columnDefinitionsSql = File.ReadAllText(columnDefinitionsSqlPath);

        var outputLines = new List<string>();
        var selectLines = new List<string>();
        var deleteLines = new List<string>();
        var insertLines = new List<string>();
        var flashbackLines = new List<string>();
        var logLines = new List<string>();

        using var connection = new OracleConnection(connectionString);
        connection.Open();

        if (!string.IsNullOrWhiteSpace(preSql))
        {
            using var preCommand = connection.CreateCommand();
            preCommand.CommandText = preSql;
            preCommand.CommandType = System.Data.CommandType.Text;
            var outParam = preCommand.Parameters.Add("result", OracleDbType.Varchar2, 4000);
            outParam.Direction = System.Data.ParameterDirection.Output;
            try
            {
                preCommand.ExecuteNonQuery();
                logLines.Add($"-- PreSql result: {outParam.Value}");
            }
            catch (Exception ex)
            {
                logLines.Add($"-- PreSql execution failed: {ex}");
            }
        }

        var primaryKeys = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        using (var primaryKeyCommand = connection.CreateCommand())
        {
            primaryKeyCommand.CommandText = primaryKeysSql;
            primaryKeyCommand.CommandType = System.Data.CommandType.Text;
            primaryKeyCommand.BindByName = true;
            primaryKeyCommand.Parameters.Add("schema", OracleDbType.Varchar2, 128).Value = schema;

            using var reader = primaryKeyCommand.ExecuteReader();
            while (reader.Read())
            {
                var table = reader.GetString(1);
                var column = reader.GetString(2);
                if (!primaryKeys.TryGetValue(table, out var list))
                {
                    list = new List<string>();
                    primaryKeys[table] = list;
                }

                list.Add(column);
            }
        }

        var tables = new Dictionary<string, List<ColumnMeta>>(StringComparer.OrdinalIgnoreCase);
        using (var columnCommand = connection.CreateCommand())
        {
            columnCommand.CommandText = columnDefinitionsSql;
            columnCommand.CommandType = System.Data.CommandType.Text;
            columnCommand.BindByName = true;
            columnCommand.Parameters.Add("schema", OracleDbType.Varchar2, 128).Value = schema;
            columnCommand.Parameters.Add("target_column", OracleDbType.Varchar2, 128).Value = targetColumn;

            using var reader = columnCommand.ExecuteReader();
            while (reader.Read())
            {
                var table = reader.GetString(1);
                var column = reader.GetString(2);
                var comment = reader.IsDBNull(3) ? string.Empty : reader.GetString(3) ?? string.Empty;
                var dataType = reader.GetString(5);

                if (!tables.TryGetValue(table, out var list))
                {
                    list = new List<ColumnMeta>();
                    tables[table] = list;
                }

                list.Add(new ColumnMeta(column, comment, dataType));
            }
        }

        var counter = 1;
        foreach (var (table, meta) in tables.OrderBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase))
        {
            var columnNames = meta.Select(m => m.Name).ToArray();
            var columnComments = meta.Select(m => m.Comment).ToArray();
            var dataTypes = meta.Select(m => m.DataType).ToArray();
            var whereClause = $"{targetColumn} {condition}";
            var orderClause = primaryKeys.TryGetValue(table, out var keys) && keys.Count > 0
                ? " ORDER BY " + string.Join(", ", keys)
                : string.Empty;

            var selectSql = $"SELECT {string.Join(", ", columnNames)} FROM {table} WHERE {whereClause}";

            using var dataCommand = connection.CreateCommand();
            dataCommand.CommandText = selectSql;
            dataCommand.CommandType = System.Data.CommandType.Text;

            using var reader = dataCommand.ExecuteReader();
            if (!reader.HasRows)
            {
                continue;
            }

            selectLines.Add($"-- {counter} {table}");
            selectLines.Add("-- * Captions " + string.Join(", ", columnComments));
            selectLines.Add(selectSql + orderClause + ";");

            deleteLines.Add($"DELETE FROM {table} WHERE {whereClause};");

            while (reader.Read())
            {
                var values = new List<string>();
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    var value = reader.GetValue(i);
                    var typeInfo = dataTypes[i];
                    values.Add(ConvertValueToSqlLiteral(value, typeInfo, logLines));
                }

                insertLines.Add($"INSERT INTO {table} ({string.Join(", ", columnNames)}) VALUES ({string.Join(", ", values)});");
            }

            if (!string.IsNullOrWhiteSpace(flashback))
            {
                flashbackLines.Add($"-- {counter} {table}");
                var flashbackSql = $"SELECT {string.Join(", ", columnNames)} FROM {table} AS OF TIMESTAMP TO_TIMESTAMP('{flashback}','YYYY-MM-DD HH24:MI:SS') WHERE {whereClause}";
                if (!string.IsNullOrWhiteSpace(orderClause))
                {
                    flashbackSql += orderClause;
                }

                flashbackLines.Add("-- * Captions " + string.Join(", ", columnComments));
                flashbackLines.Add(flashbackSql + ";");
            }

            counter++;
        }

        var sections = new List<(string Header, List<string> Lines)>
        {
            ("SELECT", selectLines),
            ("DELETE", deleteLines),
            ("INSERT", insertLines)
        };

        if (flashbackLines.Count > 0)
        {
            sections.Add(("FLASHBACK", flashbackLines));
        }

        if (logLines.Count > 0)
        {
            sections.Add(("LOG / WARNINGS", logLines));
        }

        foreach (var (header, lines) in sections)
        {
            var headerLine = $"\n-------- {header} --------";
            Console.WriteLine(headerLine);
            outputLines.Add(headerLine);
            foreach (var line in lines)
            {
                Console.WriteLine(line);
                outputLines.Add(line);
            }
        }

        if (!string.IsNullOrWhiteSpace(outputFile))
        {
            WriteOutputFiles(outputFile, outputLines, sections, condition, splitByOperation);
        }

        return 0;
    }

    private static Dictionary<string, string> ParseArgs(string[] args)
    {
        var options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (!arg.StartsWith("--", StringComparison.Ordinal))
            {
                continue;
            }

            var key = arg[2..];
            if (string.Equals(key, "split-by-operation", StringComparison.OrdinalIgnoreCase))
            {
                options[key] = "true";
                continue;
            }

            if (i + 1 >= args.Length)
            {
                continue;
            }

            options[key] = args[i + 1];
            i++;
        }

        return options;
    }

    private static string ResolveSqlDirectory(string? explicitSqlDirectory)
    {
        if (!string.IsNullOrWhiteSpace(explicitSqlDirectory))
        {
            return explicitSqlDirectory;
        }

        var checkedDirectories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var current = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        while (!string.IsNullOrWhiteSpace(current) && checkedDirectories.Add(current))
        {
            var sqlCandidate = Path.Combine(current, "sql");
            if (Directory.Exists(sqlCandidate))
            {
                return sqlCandidate;
            }

            var parent = Directory.GetParent(current);
            if (parent is null)
            {
                break;
            }

            current = parent.FullName;
        }

        throw new DirectoryNotFoundException($"SQL directory not found relative to '{AppContext.BaseDirectory}'.");
    }

    private static void WriteOutputFiles(
        string outputFile,
        List<string> outputLines,
        List<(string Header, List<string> Lines)> sections,
        string condition,
        bool splitByOperation)
    {
        var outputDirectory = Path.GetDirectoryName(outputFile);
        if (!string.IsNullOrWhiteSpace(outputDirectory) && !Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        if (!splitByOperation)
        {
            File.WriteAllLines(outputFile, outputLines, Encoding.UTF8);
            return;
        }

        var baseName = Path.GetFileNameWithoutExtension(outputFile);
        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = "query";
        }

        var conditionFragment = GetSafeFileName(condition);
        if (string.IsNullOrWhiteSpace(conditionFragment))
        {
            conditionFragment = "condition";
        }

        var groupPrefix = baseName.Contains(conditionFragment, StringComparison.OrdinalIgnoreCase)
            ? baseName
            : $"{baseName}_{conditionFragment}";

        foreach (var (header, lines) in sections)
        {
            if (lines.Count == 0)
            {
                continue;
            }

            var sectionLabel = GetSafeFileName(header.ToLowerInvariant());
            if (string.IsNullOrWhiteSpace(sectionLabel))
            {
                sectionLabel = "section";
            }

            var sectionFileName = $"{groupPrefix}_{sectionLabel}.sql";
            var sectionPath = string.IsNullOrWhiteSpace(outputDirectory)
                ? sectionFileName
                : Path.Combine(outputDirectory, sectionFileName);

            var content = new List<string> { $"-------- {header} --------" };
            content.AddRange(lines);
            File.WriteAllLines(sectionPath, content, Encoding.UTF8);
        }
    }

    private static string ConvertValueToSqlLiteral(object value, string typeInfo, List<string> logLines)
    {
        if (value is DBNull)
        {
            return "NULL";
        }

        if (typeInfo.Contains("CHAR", StringComparison.OrdinalIgnoreCase)
            || typeInfo.Contains("CLOB", StringComparison.OrdinalIgnoreCase))
        {
            var escaped = value.ToString()?.Replace("'", "''") ?? string.Empty;
            return $"'{escaped}'";
        }

        if (typeInfo.Contains("DATE", StringComparison.OrdinalIgnoreCase)
            || typeInfo.Contains("TIMESTAMP", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var dt = Convert.ToDateTime(value, InvariantCulture);
                return $"TO_DATE('{dt:yyyy-MM-dd HH:mm:ss}','YYYY-MM-DD HH24:MI:SS')";
            }
            catch (Exception ex)
            {
                logLines.Add($"-- 日付変換失敗: {ex.Message}");
                return "NULL";
            }
        }

        return Convert.ToString(value, InvariantCulture) ?? "NULL";
    }

    private static string GetSafeFileName(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var safe = System.Text.RegularExpressions.Regex.Replace(value, "\\s+", "_");
        safe = System.Text.RegularExpressions.Regex.Replace(safe, "[^0-9A-Za-z_\\-]", "_");
        safe = System.Text.RegularExpressions.Regex.Replace(safe, "_{2,}", "_");
        safe = safe.Trim('_');

        if (safe.Length > 80)
        {
            safe = safe[..80];
        }

        return safe;
    }

    private static void PrintUsage()
    {
        Console.WriteLine("SqlSdiCli - Export Oracle table data to SQL statements");
        Console.WriteLine();
        Console.WriteLine("Required:");
        Console.WriteLine("  --connection-string  Oracle connection string");
        Console.WriteLine("  --schema             Target schema name");
        Console.WriteLine("  --target-column      Target column name");
        Console.WriteLine("  --condition          WHERE clause fragment (e.g. = 'VALUE')");
        Console.WriteLine();
        Console.WriteLine("Optional:");
        Console.WriteLine("  --flashback-timestamp  Flashback timestamp (YYYY-MM-DD HH:mm:ss)");
        Console.WriteLine("  --pre-sql              PL/SQL to run before export");
        Console.WriteLine("  --sql-dir              Path to sql directory (default: search from executable)");
        Console.WriteLine("  --output-file           Output file (default: write to stdout only)");
        Console.WriteLine("  --split-by-operation    Split output by section");
        Console.WriteLine("  --help                  Show this help");
    }
}
