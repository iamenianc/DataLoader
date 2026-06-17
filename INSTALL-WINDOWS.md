# ExcelStage — Easy Windows Install (no technical knowledge needed)

There is **nothing to install** — you download a zip, unzip it, and run the
program inside. It does **not** need Excel, and it does **not** need .NET to be
installed on your computer.

---

## Step 1 — Get the program (`ExcelStage-windows.zip`)

1. Go to the project's **Releases** page on GitHub:
   `https://github.com/iamenianc/DataLoader/releases`
2. Under the newest release, click **`ExcelStage-windows.zip`** to download it.
3. In your **Downloads** folder, **right-click** the zip → **Extract All…** →
   **Extract**. This creates a folder containing `ExcelStage.exe` and the files
   it needs.
4. Move that folder somewhere easy to find, like your **Desktop**.

> Windows SmartScreen may say *"Windows protected your PC."* This is normal for
> a freshly downloaded program. Click **More info → Run anyway**.
>
> Important: keep `ExcelStage.exe` together with the other files from the zip —
> don't pull just the `.exe` out on its own, or it won't be able to connect to
> SQL Server.

*(No Releases yet? Ask whoever set this up to "publish a release," or see
"For the person who builds it" at the bottom.)*

---

## Step 2 — Put your spreadsheet in that folder (recommended)

ExcelStage automatically lists the Excel files in **whatever folder you run it
from**. The simplest setup is to copy your `.xlsx` file into the unzipped
ExcelStage folder.

---

## Step 3 — Open the Command Prompt in that folder

The easy trick:

1. Open the folder that has your spreadsheet (and `ExcelStage.exe`) in
   **File Explorer**.
2. Click once in the **address bar** at the top (where the folder path is).
3. Type `cmd` and press **Enter**.

A black Command Prompt window opens, already pointed at that folder.

---

## Step 4 — Run it

In the Command Prompt window, type:

```
ExcelStage.exe
```

…and press **Enter**. From here the program asks you simple questions:

1. **Pick your Excel file** — use the ↑ / ↓ arrow keys, press **Enter**.
2. **Pick the worksheet(s)** — arrow keys to move, **Space** to tick one or more
   sheets, then **Enter**. If you tick several, their columns are merged into one
   table and the **first sheet you tick** decides the column types.
3. **Type the SQL Server name** (for example `DBPROD-01`).
4. **Type the database name** exactly as it exists on that server.
5. **Type a name** for the table.
6. Review the summary and press **Enter** to load (or **Esc** to cancel).

That's it. It signs in to SQL Server using your normal Windows account.

---

## Quick troubleshooting

| You see…                                   | What to do                                                |
| ------------------------------------------ | --------------------------------------------------------- |
| "is not recognized as a command"           | Make sure `ExcelStage.exe` is in the folder you opened cmd in. |
| "Windows protected your PC"                | Click **More info → Run anyway**.                         |
| Your file isn't in the list                | Make sure it's a `.xlsx`/`.xlsm` in that folder (`.xlsb` won't work — re-save as `.xlsx`). |
| Login / "cannot open database" errors       | Check the server and database names are typed exactly, and that your Windows account has access to them. |
| "the 'db_upload' schema does not exist"      | Ask your DBA to create it once in that database: `CREATE SCHEMA [db_upload];` |

---

## For the person who builds it (one time)

To create `ExcelStage-windows.zip` for everyone else, on a PC with the
**.NET 8 SDK** installed (`winget install Microsoft.DotNet.SDK.8`):

- **Easiest:** in GitHub, create a tag/release named like `v1.0.0`. The
  included GitHub Action builds `ExcelStage-windows.zip` automatically and
  attaches it to that release for everyone to download.
- **Or locally:** double-click **`publish-windows.bat`** in this project. The
  finished `dist\ExcelStage-windows.zip` can be handed to any Windows PC.

> The build is packaged as a zipped folder (not a single `.exe`) on purpose:
> the SQL Server networking library must sit next to the `.exe`, otherwise it
> crashes when connecting to SQL Server.

> **If you have the .NET SDK, the simplest way to run it is without building an
> exe at all:** `dotnet run --project ExcelStage` from the project folder. This
> always uses the SQL networking library correctly.
