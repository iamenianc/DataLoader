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

> **Not a developer?** See **[INSTALL-WINDOWS.md](INSTALL-WINDOWS.md)** for the
> no-install, download-one-file-and-run guide.

## Running it

You don't need to remember any command-line switches — just run it from the
folder that holds your spreadsheet and answer the questions:

```
dotnet run --project ExcelStage
```

It guides you step by step:

| Step               | How you choose it                                                        |
| ------------------ | ----------------------------------------------------------------------- |
| Excel workbook     | Pick from the `.xlsx` / `.xlsm` / `.xlsb` files in the current folder (Up/Down arrows, Enter), or choose "Enter a path manually" |
| Worksheet          | Pick from the workbook's sheets with the arrow keys (auto-selected if there's only one) |
| SQL Server         | Type the server name, e.g. `DBPROD-01` or `localhost\SQLEXPRESS`         |
| Database           | Pick from the databases on that server with the arrow keys, or enter one manually |
| Staging table name | Type a name — your login is prefixed automatically                      |

In every list: **Up/Down** moves, **Enter** confirms, **Esc** cancels. Before
anything is written it shows a full confirmation summary (workbook, worksheet,
server, database, `db_upload` schema, table, column and row counts); press
**Enter** to execute or **Esc** to cancel. The destination schema is always
`db_upload`.

> `.xlsb` (binary) workbooks are listed but can't be read — re-save them as
> `.xlsx` or `.xlsm` in Excel first.

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
