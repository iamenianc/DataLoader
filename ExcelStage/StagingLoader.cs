using System.Data;
using System.Text;
using Microsoft.Data.SqlClient;

namespace ExcelStage;

/// <summary>
/// Creates the staging table under the <c>db_upload</c> schema and bulk-loads
/// the worksheet rows into it via SqlBulkCopy.
/// </summary>
public static class StagingLoader
{
    // The staging schema is fixed and must NEVER be user-supplied or overridden.
    // It is the single source of truth for every schema-qualified statement below.
    public const string SchemaName = "db_upload";

    public static int Load(
        string connectionString,
        string tableName,
        IReadOnlyList<InferredColumn> columns,
        ExcelSheet sheet)
    {
        // Safety guard: refuse to run if the fixed schema is ever anything other
        // than "db_upload" (e.g. a future edit accidentally changes the constant).
        if (!string.Equals(SchemaName, "db_upload", StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"Refusing to run: the staging schema must be 'db_upload' but was '{SchemaName}'.");
        }

        using var connection = new SqlConnection(connectionString);
        connection.Open();

        EnsureSchema(connection);
        CreateTable(connection, tableName, columns);

        var table = BuildDataTable(columns, sheet);
        BulkCopy(connection, tableName, columns, table);

        return table.Rows.Count;
    }

    private static void EnsureSchema(SqlConnection connection)
    {
        const string sql = """
            IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = @schema)
                EXEC('CREATE SCHEMA [db_upload]');
            """;

        using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@schema", SchemaName);
        command.ExecuteNonQuery();
    }

    private static void CreateTable(SqlConnection connection, string tableName, IReadOnlyList<InferredColumn> columns)
    {
        var qualified = $"[{SchemaName}].[{tableName}]";

        // Drop any existing staging table so re-runs start clean.
        var dropSql = $"IF OBJECT_ID(N'{qualified}', N'U') IS NOT NULL DROP TABLE {qualified};";
        using (var drop = new SqlCommand(dropSql, connection))
        {
            drop.ExecuteNonQuery();
        }

        var builder = new StringBuilder();
        builder.Append("CREATE TABLE ").Append(qualified).AppendLine(" (");
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            builder.Append("    [").Append(column.Name).Append("] ").Append(column.SqlType);
            builder.Append(column.IsNullable ? " NULL" : " NOT NULL");
            builder.AppendLine(i < columns.Count - 1 ? "," : string.Empty);
        }

        builder.AppendLine(");");

        var createSql = builder.ToString();
        Console.WriteLine(createSql);

        using var create = new SqlCommand(createSql, connection);
        create.ExecuteNonQuery();
    }

    private static DataTable BuildDataTable(IReadOnlyList<InferredColumn> columns, ExcelSheet sheet)
    {
        var table = new DataTable();
        foreach (var column in columns)
        {
            table.Columns.Add(new DataColumn(column.Name, ClrType(column.DbType)) { AllowDBNull = true });
        }

        foreach (var row in sheet.Rows)
        {
            var values = new object?[columns.Count];
            for (var c = 0; c < columns.Count; c++)
            {
                var cell = c < row.Length ? row[c] : ExcelCell.Empty;
                values[c] = columns[c].Convert(cell);
            }

            table.Rows.Add(values);
        }

        return table;
    }

    private static void BulkCopy(
        SqlConnection connection, string tableName, IReadOnlyList<InferredColumn> columns, DataTable table)
    {
        using var bulkCopy = new SqlBulkCopy(connection)
        {
            DestinationTableName = $"[{SchemaName}].[{tableName}]",
            BulkCopyTimeout = 0,
            BatchSize = 5000
        };

        foreach (var column in columns)
        {
            bulkCopy.ColumnMappings.Add(column.Name, column.Name);
        }

        bulkCopy.WriteToServer(table);
    }

    private static Type ClrType(SqlDbType dbType) => dbType switch
    {
        SqlDbType.Bit => typeof(bool),
        SqlDbType.BigInt => typeof(long),
        SqlDbType.Decimal => typeof(decimal),
        SqlDbType.Float => typeof(double),
        SqlDbType.DateTime2 => typeof(DateTime),
        _ => typeof(string)
    };
}
