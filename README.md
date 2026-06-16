# ExcelStage

A small, single-purpose .NET 8 console tool that imports an Excel worksheet
into a SQL Server **staging table** using `SqlBulkCopy`.

It deliberately does only one thing:

1. Reads an `.xlsx` workbook (OpenXML, no Excel/Office required).
2. Reads the first worksheet, or one you name.
3. Infers a SQL column type for each column from the data.
4. Creates the table in the `db_upload` schema (dropping any old copy first).
5. Bulk-loads every row with `SqlBulkCopy`.

The table is created as `db_upload.<your-login>_<name-you-choose>` so each
person's staging tables stay separate.

## Running it

You don't need to remember any command-line switches — just run it and
answer the questions:

```
dotnet run --project ExcelStage
```

It will prompt you for:

| Prompt              | Example                     | Notes                                   |
| ------------------- | --------------------------- | --------------------------------------- |
| Excel workbook path | `C:\data\customers.xlsx`    | Must be an existing `.xlsx` file        |
| Worksheet name      | `Sheet1`                    | Leave blank for the first worksheet     |
| SQL Server          | `DBPROD-01`                 | Server name or `host\instance`          |
| Database name       | `Sales`                     | Must already exist                      |
| Staging table name  | `customers`                 | Your login is prefixed automatically    |

Before anything is written it shows a summary and asks you to confirm.

## Connection

ExcelStage connects with Windows authentication (Integrated Security) using
an encrypted connection, equivalent to:

```
Data Source=<server>;Initial Catalog=<database>;Integrated Security=True;
Persist Security Info=False;Pooling=False;MultipleActiveResultSets=False;
Encrypt=True;TrustServerCertificate=True;
```

## Type inference

For each column ExcelStage looks at every value in the sheet and picks the
narrowest type that fits:

| If every value is…              | Column type        |
| ------------------------------- | ------------------ |
| a whole number (incl. 0/1)      | `BIGINT`           |
| true/false or yes/no            | `BIT`              |
| a number with decimals          | `DECIMAL(38, 10)`  |
| a date/time                     | `DATETIME2`        |
| anything else                   | `NVARCHAR(n)`      |

Mixed or empty columns fall back to `NVARCHAR`. The generated `CREATE TABLE`
statement is printed to the console before it runs.

## Build

```
dotnet build ExcelStage
```

Requires the .NET 8 SDK. NuGet dependencies:

- `DocumentFormat.OpenXml` — read the `.xlsx`
- `Microsoft.Data.SqlClient` — `CREATE TABLE` + `SqlBulkCopy`
