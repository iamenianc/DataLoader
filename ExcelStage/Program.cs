using System.Text;
using ExcelStage;
using Microsoft.Data.SqlClient;

Console.OutputEncoding = Encoding.UTF8;

PrintBanner();

try
{
    // ----- 1. Excel workbook -------------------------------------------------
    var excelPath = PromptForExistingFile(
        "Path to the Excel workbook to import (.xlsx)",
        "e.g. C:\\data\\customers.xlsx");

    var worksheet = PromptOptional(
        "Worksheet name to import",
        "Leave blank to use the first worksheet in the workbook");

    Console.WriteLine();
    Console.WriteLine($"Reading '{Path.GetFileName(excelPath)}'...");
    var sheet = ExcelReader.Read(excelPath, worksheet);
    Console.WriteLine($"  Found {sheet.Headers.Count} column(s) and {sheet.Rows.Count} data row(s).");

    Console.WriteLine("Inferring SQL column types from the data...");
    var columns = ColumnTypeInferrer.Infer(sheet);
    PrintColumnSummary(columns);

    // ----- 2. SQL Server target ---------------------------------------------
    Console.WriteLine();
    Console.WriteLine("Now let's tell ExcelStage where to load the data.");

    var server = PromptRequired(
        "SQL Server name / instance",
        "e.g. DBPROD-01  or  localhost\\SQLEXPRESS");

    var database = PromptRequired(
        "Database name",
        "The database that already exists on that server");

    var login = SanitizeLogin(Environment.UserName);
    var baseName = PromptRequired(
        "Staging table name",
        $"Your login ('{login}') will be added automatically so the table is yours");

    var tableName = $"{login}_{SanitizeLogin(baseName)}";
    var connectionString = BuildConnectionString(server, database);

    // ----- 3. Confirm before touching the database --------------------------
    Console.WriteLine();
    Console.WriteLine("About to run with these settings:");
    Console.WriteLine($"  Server      : {server}");
    Console.WriteLine($"  Database    : {database}");
    Console.WriteLine($"  Destination : [{StagingLoader.SchemaName}].[{tableName}]");
    Console.WriteLine($"  Rows        : {sheet.Rows.Count}");
    Console.WriteLine();
    Console.WriteLine("If the table already exists it will be dropped and recreated.");

    if (!PromptYesNo("Proceed?", defaultYes: true))
    {
        Console.WriteLine("Cancelled. No changes were made.");
        return 0;
    }

    // ----- 4. Create + load --------------------------------------------------
    Console.WriteLine();
    Console.WriteLine("Connecting and creating the staging table...");
    var inserted = StagingLoader.Load(connectionString, tableName, columns, sheet);

    Console.WriteLine();
    Console.WriteLine("Done.");
    Console.WriteLine($"  Inserted {inserted} row(s) into [{StagingLoader.SchemaName}].[{tableName}].");
    return 0;
}
catch (FileNotFoundException ex)
{
    return Fail($"File not found: {ex.Message}");
}
catch (InvalidDataException ex)
{
    return Fail($"There was a problem reading the workbook: {ex.Message}");
}
catch (SqlException ex)
{
    return Fail($"SQL Server reported an error: {ex.Message}");
}
catch (Exception ex)
{
    return Fail(ex.Message);
}

// ---------------------------------------------------------------------------
// Console helpers
// ---------------------------------------------------------------------------

static void PrintBanner()
{
    Console.WriteLine("======================================================");
    Console.WriteLine(" ExcelStage - import an Excel sheet into SQL Server");
    Console.WriteLine("======================================================");
    Console.WriteLine("This tool will ask you a few questions, read your");
    Console.WriteLine("spreadsheet, build a matching table, and load the rows.");
    Console.WriteLine("Press Ctrl+C at any time to cancel.");
    Console.WriteLine();
}

static void PrintColumnSummary(List<InferredColumn> columns)
{
    Console.WriteLine();
    Console.WriteLine("  Inferred schema:");
    foreach (var column in columns)
    {
        Console.WriteLine($"    - [{column.Name}] {column.SqlType}");
    }
}

static string PromptForExistingFile(string label, string hint)
{
    while (true)
    {
        var value = PromptRequired(label, hint);
        value = value.Trim().Trim('"');
        if (File.Exists(value))
        {
            return value;
        }

        Console.WriteLine($"  ! No file found at '{value}'. Please try again.");
    }
}

static string PromptRequired(string label, string hint)
{
    while (true)
    {
        Console.WriteLine();
        if (!string.IsNullOrEmpty(hint))
        {
            Console.WriteLine($"  ({hint})");
        }

        Console.Write($"{label}: ");
        var value = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(value))
        {
            return value.Trim();
        }

        Console.WriteLine("  ! This value is required. Please enter something.");
    }
}

static string? PromptOptional(string label, string hint)
{
    Console.WriteLine();
    if (!string.IsNullOrEmpty(hint))
    {
        Console.WriteLine($"  ({hint})");
    }

    Console.Write($"{label}: ");
    var value = Console.ReadLine();
    return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}

static bool PromptYesNo(string label, bool defaultYes)
{
    var suffix = defaultYes ? "[Y/n]" : "[y/N]";
    while (true)
    {
        Console.Write($"{label} {suffix}: ");
        var value = Console.ReadLine()?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(value))
        {
            return defaultYes;
        }

        if (value is "y" or "yes") return true;
        if (value is "n" or "no") return false;

        Console.WriteLine("  ! Please answer y or n.");
    }
}

static string SanitizeLogin(string raw)
{
    // Strip a DOMAIN\ prefix and keep only identifier-safe characters.
    var name = raw.Contains('\\') ? raw[(raw.LastIndexOf('\\') + 1)..] : raw;
    var chars = name.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray();
    var result = new string(chars).Trim('_');
    if (string.IsNullOrEmpty(result))
    {
        result = "user";
    }

    return char.IsDigit(result[0]) ? "_" + result : result;
}

static string BuildConnectionString(string server, string database) =>
    new SqlConnectionStringBuilder
    {
        DataSource = server,
        InitialCatalog = database,
        IntegratedSecurity = true,
        PersistSecurityInfo = false,
        Pooling = false,
        MultipleActiveResultSets = false,
        Encrypt = true,
        TrustServerCertificate = true
    }.ConnectionString;

static int Fail(string message)
{
    Console.Error.WriteLine();
    Console.Error.WriteLine($"ERROR: {message}");
    return 1;
}
