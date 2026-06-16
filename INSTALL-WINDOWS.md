# ExcelStage — Easy Windows Install (no technical knowledge needed)

ExcelStage is a single program file. There is **nothing to install** — you
download one file and run it. It does **not** need Excel, and it does **not**
need .NET to be installed on your computer.

---

## Step 1 — Get the program (`ExcelStage.exe`)

1. Go to the project's **Releases** page on GitHub:
   `https://github.com/iamenianc/DataLoader/releases`
2. Under the newest release, click **`ExcelStage.exe`** to download it.
3. Save it somewhere easy to find, like your **Desktop** or your
   **Documents** folder.

> Windows SmartScreen may say *"Windows protected your PC."* This is normal for
> a freshly downloaded program. Click **More info → Run anyway**.

*(No Releases yet? Ask whoever set this up to "publish a release," or see
"For the person who builds it" at the bottom.)*

---

## Step 2 — Put it next to your spreadsheets (recommended)

ExcelStage automatically lists the Excel files in **whatever folder you run it
from**. The simplest setup is to copy `ExcelStage.exe` into the same folder as
the `.xlsx` file you want to load.

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
2. **Pick the worksheet** — arrow keys, **Enter**.
3. **Type the SQL Server name** (for example `DBPROD-01`).
4. **Pick the database** — arrow keys, **Enter**.
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
| "Could not list databases" / login errors  | Your Windows account needs access to that SQL Server / database. |

---

## For the person who builds it (one time)

To create `ExcelStage.exe` for everyone else, on a PC with the
**.NET 8 SDK** installed (`winget install Microsoft.DotNet.SDK.8`):

- **Easiest:** in GitHub, create a tag/release named like `v1.0.0`. The
  included GitHub Action builds `ExcelStage.exe` automatically and attaches it
  to that release for everyone to download.
- **Or locally:** double-click **`publish-windows.bat`** in this project. The
  finished `dist\ExcelStage.exe` can be copied to any Windows PC.
