using System.Text;
using ExcelStage;
using Microsoft.Data.SqlClient;

Console.OutputEncoding = Encoding.UTF8;

PrintBanner();

try
{
    string? excelPath = null;
    string? worksheet = null;
    string? server = null;
    string? database = null;
    string? baseName = null;
    ExcelSheet? sheet = null;
    List<InferredColumn>? columns = null;
    var login = SanitizeLogin(Environment.UserName);

    var step = Step.File;
    while (true)
    {
        switch (step)
        {
            case Step.File:
            {
                var r = SelectExcelFile();
                if (r.Kind is NavKind.Quit or NavKind.Back) return Cancelled();
                if (r.Kind == NavKind.Restart) { step = Step.File; continue; }
                excelPath = r.Value;
                step = Step.Worksheet;
                continue;
            }

            case Step.Worksheet:
            {
                var r = SelectWorksheet(excelPath!);
                if (r.Kind == NavKind.Quit) return Cancelled();
                if (r.Kind is NavKind.Restart or NavKind.Back) { step = Step.File; continue; }
                worksheet = r.Value;

                Console.WriteLine();
                Console.WriteLine($"Reading '{Path.GetFileName(excelPath)}' / '{worksheet}'...");
                try
                {
                    sheet = ExcelReader.Read(excelPath!, worksheet!);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  ! Could not read that worksheet: {ex.Message}");
                    Console.WriteLine("    Pick another worksheet, or press B to choose a different file.");
                    continue; // stay on the worksheet step
                }

                Console.WriteLine($"  Found {sheet.Headers.Count} column(s) and {sheet.Rows.Count} data row(s).");
                Console.WriteLine("Inferring SQL column types from the data...");
                columns = ColumnTypeInferrer.Infer(sheet);
                PrintColumnSummary(columns);
                step = Step.Server;
                continue;
            }

            case Step.Server:
            {
                var r = PromptNav("SQL Server name / instance", "e.g. DBPROD-01  or  localhost\\SQLEXPRESS");
                if (r.Kind == NavKind.Quit) return Cancelled();
                if (r.Kind == NavKind.Restart) { step = Step.File; continue; }
                if (r.Kind == NavKind.Back) { step = Step.Worksheet; continue; }
                server = r.Value;
                step = Step.Database;
                continue;
            }

            case Step.Database:
            {
                Console.WriteLine();
                Console.WriteLine($"Connecting to '{server}' to list databases...");
                var r = SelectDatabase(server!);
                if (r.Kind == NavKind.Quit) return Cancelled();
                if (r.Kind == NavKind.Restart) { step = Step.File; continue; }
                if (r.Kind == NavKind.Back) { step = Step.Server; continue; }
                database = r.Value;
                step = Step.Table;
                continue;
            }

            case Step.Table:
            {
                var r = PromptNav("Staging table name",
                    $"Your login ('{login}') will be added automatically so the table is yours");
                if (r.Kind == NavKind.Quit) return Cancelled();
                if (r.Kind == NavKind.Restart) { step = Step.File; continue; }
                if (r.Kind == NavKind.Back) { step = Step.Database; continue; }
                baseName = r.Value;
                step = Step.Confirm;
                continue;
            }

            case Step.Confirm:
            {
                var tableName = $"{login}_{SanitizeLogin(baseName!)}";
                PrintSummary(excelPath!, worksheet!, server!, database!, tableName, columns!.Count, sheet!.Rows.Count);

                var r = ConfirmNav();
                if (r.Kind == NavKind.Quit) return Cancelled();
                if (r.Kind == NavKind.Restart) { step = Step.File; continue; }
                if (r.Kind == NavKind.Back) { step = Step.Table; continue; }

                var connectionString = Sql.BuildConnectionString(server!, database!);
                Console.WriteLine();
                Console.WriteLine("Connecting and creating the staging table...");
                var inserted = StagingLoader.Load(connectionString, tableName, columns!, sheet!);

                Console.WriteLine();
                Console.WriteLine("Done.");
                Console.WriteLine($"  Inserted {inserted} row(s) into [{StagingLoader.SchemaName}].[{tableName}].");
                return 0;
            }
        }
    }
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
// Selection flows (each returns a value or a navigation request)
// ---------------------------------------------------------------------------

static Nav<string> SelectExcelFile()
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
        if (!choice.IsValue)
        {
            return choice.Carry<string>();
        }

        if (choice.Value == options.Count - 1)
        {
            var manual = PromptForExistingFileNav("Full path to the Excel workbook", "e.g. C:\\data\\customers.xlsx");
            if (manual.Kind == NavKind.Back)
            {
                continue; // back to the file list
            }

            if (!manual.IsValue)
            {
                return manual; // restart / quit
            }

            if (IsBinaryWorkbook(manual.Value!))
            {
                WarnBinaryWorkbook();
                continue;
            }

            return manual;
        }

        var selected = files[choice.Value];
        if (IsBinaryWorkbook(selected))
        {
            WarnBinaryWorkbook();
            continue;
        }

        return Nav<string>.FromValue(selected);
    }
}

static Nav<string> SelectWorksheet(string path)
{
    var names = ExcelReader.GetWorksheetNames(path);
    if (names.Count == 0)
    {
        throw new InvalidDataException("The workbook does not contain any worksheets.");
    }

    var choice = Menu.Select("Select a worksheet:", names);
    return choice.IsValue ? Nav<string>.FromValue(names[choice.Value]) : choice.Carry<string>();
}

static Nav<string> SelectDatabase(string server)
{
    List<string> databases;
    try
    {
        databases = Sql.ListDatabases(server);
    }
    catch (Exception ex)
    {
        Console.WriteLine("  ! Could not list databases automatically.");
        Console.WriteLine($"    Reason: {Describe(ex)}");
        Console.WriteLine("    You can still type the database name to continue.");
        return PromptNav("Database name", "Type the database name to use");
    }

    if (databases.Count == 0)
    {
        Console.WriteLine("  ! No databases were returned for your login.");
        return PromptNav("Database name", "Type the database name to use");
    }

    var options = new List<string>(databases) { "Enter a name manually..." };
    while (true)
    {
        var choice = Menu.Select("Select a database:", options);
        if (!choice.IsValue)
        {
            return choice.Carry<string>();
        }

        if (choice.Value == options.Count - 1)
        {
            var manual = PromptNav("Database name", "Type the database name to use");
            if (manual.Kind == NavKind.Back)
            {
                continue; // back to the database list
            }

            return manual;
        }

        return Nav<string>.FromValue(databases[choice.Value]);
    }
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
// Prompts and console helpers
// ---------------------------------------------------------------------------

static void PrintBanner()
{
    Console.WriteLine("======================================================");
    Console.WriteLine(" ExcelStage - import an Excel sheet into SQL Server");
    Console.WriteLine("======================================================");
    Console.WriteLine("This tool will ask you a few questions, read your");
    Console.WriteLine("spreadsheet, build a matching table, and load the rows.");
    Console.WriteLine();
    Console.WriteLine("At any step:  B = back   R = restart   Q / Esc = cancel");
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

static void PrintSummary(
    string excelPath, string worksheet, string server, string database,
    string tableName, int columnCount, int rowCount)
{
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
    Console.WriteLine($"  Columns     : {columnCount}");
    Console.WriteLine($"  Rows        : {rowCount}");
    Console.WriteLine("------------------------------------------------------");
    Console.WriteLine("If the table already exists it will be dropped and recreated.");
    Console.WriteLine();
}

// Asks for a required value. Typing B/R/Q (or back/restart/quit/cancel)
// navigates instead of returning a value.
static Nav<string> PromptNav(string label, string hint)
{
    while (true)
    {
        Console.WriteLine();
        if (!string.IsNullOrEmpty(hint))
        {
            Console.WriteLine($"  ({hint})");
        }

        Console.WriteLine("  (or type B = back, R = restart, Q = cancel)");
        Console.Write($"{label}: ");
        var input = Console.ReadLine();

        var command = AsCommand(input);
        if (command.HasValue)
        {
            return command.Value;
        }

        if (!string.IsNullOrWhiteSpace(input))
        {
            return Nav<string>.FromValue(input.Trim());
        }

        Console.WriteLine("  ! This value is required. Please enter something (or B/R/Q).");
    }
}

static Nav<string> PromptForExistingFileNav(string label, string hint)
{
    while (true)
    {
        var r = PromptNav(label, hint);
        if (!r.IsValue)
        {
            return r;
        }

        var path = CleanPath(r.Value!);
        if (File.Exists(path))
        {
            return Nav<string>.FromValue(path);
        }

        Console.WriteLine($"  ! No file found at '{path}'. Please try again.");
    }
}

// Recognises navigation commands typed at a prompt. Returns null for a normal value.
static Nav<string>? AsCommand(string? input)
{
    switch (input?.Trim().ToLowerInvariant())
    {
        case "b" or "back":
            return Nav<string>.Back;
        case "r" or "restart":
            return Nav<string>.Restart;
        case "q" or "quit" or "cancel" or "exit":
            return Nav<string>.Quit;
        default:
            return null;
    }
}

// Final go/no-go gate. Enter executes; B/R/Q (or Esc) navigate.
static Nav<bool> ConfirmNav()
{
    Console.Write("Press ENTER to execute  |  B = back, R = restart, Esc/Q = cancel: ");

    if (Console.IsInputRedirected)
    {
        var line = Console.ReadLine()?.Trim().ToLowerInvariant();
        Console.WriteLine();
        return line switch
        {
            "b" or "back" => Nav<bool>.Back,
            "r" or "restart" => Nav<bool>.Restart,
            "q" or "quit" or "cancel" => Nav<bool>.Quit,
            _ => Nav<bool>.FromValue(true) // blank / Enter
        };
    }

    while (true)
    {
        var key = Console.ReadKey(intercept: true).Key;
        switch (key)
        {
            case ConsoleKey.Enter:
                Console.WriteLine();
                return Nav<bool>.FromValue(true);
            case ConsoleKey.B:
                Console.WriteLine();
                return Nav<bool>.Back;
            case ConsoleKey.R:
                Console.WriteLine();
                return Nav<bool>.Restart;
            case ConsoleKey.Escape:
            case ConsoleKey.Q:
                Console.WriteLine();
                return Nav<bool>.Quit;
        }
    }
}

// Builds a readable reason from an exception chain, falling back to the type
// name when a message is empty (so the reason line is never blank).
static string Describe(Exception ex)
{
    var parts = new List<string>();
    Exception? current = ex;
    while (current is not null)
    {
        var message = string.IsNullOrWhiteSpace(current.Message)
            ? current.GetType().Name
            : current.Message.Trim();
        parts.Add(message);
        current = current.InnerException;
    }

    return string.Join("  ->  ", parts);
}

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

static int Cancelled()
{
    Console.WriteLine();
    Console.WriteLine("Cancelled. No changes were made.");
    return 0;
}

static int Fail(string message)
{
    Console.Error.WriteLine();
    Console.Error.WriteLine($"ERROR: {message}");
    return 1;
}

enum Step
{
    File,
    Worksheet,
    Server,
    Database,
    Table,
    Confirm
}
