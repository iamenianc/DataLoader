# ExcelStage

A small, single-purpose .NET 8 console tool that imports an Excel worksheet
into a SQL Server **staging table** using `SqlBulkCopy`.

It deliberately does only one thing:

1. Reads an `.xlsx` workbook (OpenXML, no Excel/Office required).
2. Reads one or more worksheets you select from that workbook.
3. Infers a SQL column type for each column from the data.
4. Creates the table in the `db_upload` schema (dropping any old copy first).
5. Bulk-loads every row with `SqlBulkCopy`.

When you pick several sheets, their columns are **unioned by name** — a column
that appears in more than one sheet becomes a single table column, and every
sheet's rows are loaded into that one table (cells are `NULL` where a sheet
doesn't have a column). Each column's **type is inferred from the first sheet
you select** that contains it.

The table is created as `db_upload.tmp_<your-login>_<name-you-choose>` so each
person's staging tables stay separate. The `db_upload` schema is fixed and can
never be changed; it must already exist in the target database (the tool
verifies this and tells you if it's missing rather than trying to create it).

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
| Worksheet(s)       | Tick one or more sheets (**Space** toggles, **Enter** confirms; ticks are numbered in pick order, and **#1 sets the column types**). Auto-selected if there's only one |
| SQL Server         | Type the exact server name, e.g. `DBPROD-01` or `localhost\SQLEXPRESS`   |
| Database           | Type the exact database name as it exists on that server                 |
| Staging table name | Type a name — it's prefixed with `tmp_` and your login automatically     |

In every list: **Up/Down** moves, **Enter** confirms, **Esc** cancels (the
worksheet list also uses **Space** to tick multiple sheets). Before
anything is written it shows a full confirmation summary (workbook, worksheet(s),
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
