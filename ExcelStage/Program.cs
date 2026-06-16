using System.Text;
using ExcelStage;
using Microsoft.Data.SqlClient;

Console.OutputEncoding = Encoding.UTF8;

PrintBanner();

try
{
    // ----- 1. Excel workbook -------------------------------------------------
    var excelPath = SelectExcelFile();
    if (excelPath is null)
    {
        Console.WriteLine("Cancelled. No file selected.");
        return 0;
    }

    // ----- 2. Worksheet ------------------------------------------------------
    var worksheet = SelectWorksheet(excelPath);
    if (worksheet is null)
    {
        Console.WriteLine("Cancelled. No worksheet selected.");
        return 0;
    }

    Console.WriteLine();
    Console.WriteLine($"Reading '{Path.GetFileName(excelPath)}' / '{worksheet}'...");
    var sheet = ExcelReader.Read(excelPath, worksheet);
    Console.WriteLine($"  Found {sheet.Headers.Count} column(s) and {sheet.Rows.Count} data row(s).");

    Console.WriteLine("Inferring SQL column types from the data...");
    var columns = ColumnTypeInferrer.Infer(sheet);
    PrintColumnSummary(columns);

    // ----- 3. SQL Server target ---------------------------------------------
    Console.WriteLine();
    Console.WriteLine("Now let's tell ExcelStage where to load the data.");

    var server = PromptRequired(
        "SQL Server name / instance",
        "e.g. DBPROD-01  or  localhost\\SQLEXPRESS");

    var database = SelectDatabase(server);
    if (database is null)
    {
        Console.WriteLine("Cancelled. No database selected.");
        return 0;
    }

    var login = SanitizeLogin(Environment.UserName);
    var baseName = PromptRequired(
        "Staging table name",
        $"Your login ('{login}') will be added automatically so the table is yours");

    var tableName = $"{login}_{SanitizeLogin(baseName)}";
    var connectionString = Sql.BuildConnectionString(server, database);

    // ----- 4. Confirm before touching the database --------------------------
    Console.WriteLine();
    Console.WriteLine("======================================================");
    Console.WriteLine(" Please confirm - about to commit the following:");
    Console.WriteLine("======================================================");
    Console.WriteLine($"  Workbook    : {Path.GetFileName(excelPath)}");
    Console.WriteLine($"  Worksheet   : {worksheet}");
    Console.WriteLine($"  Server      : {server}");
    Console.WriteLine($"  Database    : {database}");
    Console.WriteLine($"  Schema      : {StagingLoader.SchemaName}");
    Console.WriteLine($"  Table       : {tableName}");
    Console.WriteLine($"  Destination : [{StagingLoader.SchemaName}].[{tableName}]");
    Console.WriteLine($"  Columns     : {columns.Count}");
    Console.WriteLine($"  Rows        : {sheet.Rows.Count}");
    Console.WriteLine("------------------------------------------------------");
    Console.WriteLine("If the table already exists it will be dropped and recreated.");
    Console.WriteLine();

    if (!ConfirmExecute())
    {
        Console.WriteLine("Cancelled. No changes were made.");
        return 0;
    }

    // ----- 5. Create + load --------------------------------------------------
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
// Selection flows
// ---------------------------------------------------------------------------

// Returns the chosen workbook path, or null if the user cancelled.
static string? SelectExcelFile()
{
    var directory = Directory.GetCurrentDirectory();

    while (true)
    {
        var files = Directory.EnumerateFiles(directory)
            .Where(IsExcelWorkbook)
            .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        Console.WriteLine();
        Console.WriteLine($"Looking for Excel files in: {directory}");
        if (files.Count == 0)
        {
            Console.WriteLine("  (no .xlsx / .xlsm / .xlsb files found here)");
        }

        var options = files.Select(f => Path.GetFileName(f)!).ToList();
        options.Add("Enter a path manually...");

        var choice = Menu.Select("Select an Excel workbook:", options);
        if (choice < 0)
        {
            return null; // Esc
        }

        if (choice == options.Count - 1)
        {
            var manual = PromptForExistingFile(
                "Full path to the Excel workbook",
                "e.g. C:\\data\\customers.xlsx");
            if (IsBinaryWorkbook(manual))
            {
                WarnBinaryWorkbook();
                continue;
            }

            return manual;
        }

        var selected = files[choice];
        if (IsBinaryWorkbook(selected))
        {
            WarnBinaryWorkbook();
            continue;
        }

        return selected;
    }
}

// Returns the chosen worksheet name, or null if cancelled.
static string? SelectWorksheet(string path)
{
    var names = ExcelReader.GetWorksheetNames(path);
    if (names.Count == 0)
    {
        throw new InvalidDataException("The workbook does not contain any worksheets.");
    }

    if (names.Count == 1)
    {
        Console.WriteLine();
        Console.WriteLine($"Using the only worksheet: '{names[0]}'.");
        return names[0];
    }

    var choice = Menu.Select("Select a worksheet:", names);
    return choice < 0 ? null : names[choice];
}

// Returns the chosen database name, or null if cancelled.
static string? SelectDatabase(string server)
{
    List<string> databases;
    try
    {
        Console.WriteLine();
        Console.WriteLine($"Connecting to '{server}' to list databases...");
        databases = Sql.ListDatabases(server);
    }
    catch (Exception ex)
    {
        // Listing databases is a convenience only - never let it abort the run.
        // Any failure (network, login, TLS, etc.) just falls back to typing the name.
        Console.WriteLine($"  ! Could not list databases automatically.");
        Console.WriteLine($"    Reason: {ex.Message.Trim()}");
        Console.WriteLine("    You can still type the database name to continue.");
        return PromptRequired("Database name", "Type the database name to use");
    }

    if (databases.Count == 0)
    {
        Console.WriteLine("  ! No user databases were returned.");
        return PromptRequired("Database name", "Type the database name to use");
    }

    var options = new List<string>(databases) { "Enter a name manually..." };
    var choice = Menu.Select("Select a database:", options);
    if (choice < 0)
    {
        return null;
    }

    return choice == options.Count - 1
        ? PromptRequired("Database name", "Type the database name to use")
        : databases[choice];
}

static bool IsExcelWorkbook(string path)
{
    var ext = Path.GetExtension(path).ToLowerInvariant();
    return ext is ".xlsx" or ".xlsm" or ".xlsb";
}

static bool IsBinaryWorkbook(string path) =>
    Path.GetExtension(path).ToLowerInvariant() == ".xlsb";

static void WarnBinaryWorkbook()
{
    Console.WriteLine();
    Console.WriteLine("  ! .xlsb (binary) workbooks can't be read by this tool.");
    Console.WriteLine("    Please open it in Excel and 'Save As' .xlsx or .xlsm, then try again.");
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
        var value = CleanPath(PromptRequired(label, hint));
        if (File.Exists(value))
        {
            return value;
        }

        Console.WriteLine($"  ! No file found at '{value}'. Please try again.");
    }
}

// Removes surrounding single/double quotes and whitespace from a pasted path.
// Windows' "Copy as path" wraps the path in double quotes, which break File.Exists.
static string CleanPath(string raw)
{
    var value = raw.Trim();
    while (value.Length >= 2 &&
           ((value[0] == '"' && value[^1] == '"') || (value[0] == '\'' && value[^1] == '\'')))
    {
        value = value[1..^1].Trim();
    }

    return value.Trim('"', '\'').Trim();
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

// Final go/no-go gate: Enter executes, Esc (or anything else when redirected)
// cancels. Returns true to proceed.
static bool ConfirmExecute()
{
    Console.Write("Press ENTER to execute, or Esc to cancel: ");

    if (Console.IsInputRedirected)
    {
        // No interactive key access: a blank line means "go", anything else cancels.
        var line = Console.ReadLine();
        Console.WriteLine();
        return string.IsNullOrWhiteSpace(line);
    }

    while (true)
    {
        var key = Console.ReadKey(intercept: true).Key;
        if (key == ConsoleKey.Enter)
        {
            Console.WriteLine();
            return true;
        }

        if (key == ConsoleKey.Escape)
        {
            Console.WriteLine();
            return false;
        }
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

static int Fail(string message)
{
    Console.Error.WriteLine();
    Console.Error.WriteLine($"ERROR: {message}");
    return 1;
}
