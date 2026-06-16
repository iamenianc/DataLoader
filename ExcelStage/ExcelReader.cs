using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStage;

/// <summary>
/// Reads a single worksheet from an .xlsx workbook into headers + rows using
/// the OpenXML SDK. The first non-empty row is treated as the header row.
/// </summary>
public sealed class ExcelSheet
{
    public required List<string> Headers { get; init; }
    public required List<ExcelCell[]> Rows { get; init; }
}

public static class ExcelReader
{
    /// <summary>Returns the worksheet names in the workbook, in tab order.</summary>
    public static List<string> GetWorksheetNames(string path)
    {
        using var document = SpreadsheetDocument.Open(path, isEditable: false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("Workbook part is missing; the file is not a valid Excel workbook.");

        return workbookPart.Workbook.Sheets?.Elements<Sheet>()
            .Select(s => s.Name?.Value ?? string.Empty)
            .Where(n => !string.IsNullOrEmpty(n))
            .ToList() ?? new List<string>();
    }

    public static ExcelSheet Read(string path, string? worksheetName)
    {
        using var document = SpreadsheetDocument.Open(path, isEditable: false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("Workbook part is missing; the file is not a valid .xlsx workbook.");

        var sheet = ResolveSheet(workbookPart, worksheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);

        var sharedStrings = LoadSharedStrings(workbookPart);
        var dateFormatStyles = LoadDateFormatStyles(workbookPart);

        var rows = new List<ExcelCell[]>();
        var maxColumn = 0;

        foreach (var row in worksheetPart.Worksheet.Descendants<Row>())
        {
            var cellsByColumn = new Dictionary<int, ExcelCell>();
            foreach (var cell in row.Elements<Cell>())
            {
                var columnIndex = ColumnIndexFromReference(cell.CellReference?.Value);
                if (columnIndex < 0)
                {
                    continue;
                }

                cellsByColumn[columnIndex] = ParseCell(cell, sharedStrings, dateFormatStyles);
                if (columnIndex + 1 > maxColumn)
                {
                    maxColumn = columnIndex + 1;
                }
            }

            if (cellsByColumn.Count == 0)
            {
                rows.Add(Array.Empty<ExcelCell>());
                continue;
            }

            var ordered = new ExcelCell[maxColumn];
            for (var i = 0; i < maxColumn; i++)
            {
                ordered[i] = cellsByColumn.TryGetValue(i, out var c) ? c : ExcelCell.Empty;
            }

            rows.Add(ordered);
        }

        return BuildSheet(rows, maxColumn);
    }

    private static ExcelSheet BuildSheet(List<ExcelCell[]> rawRows, int maxColumn)
    {
        // The first row that contains any value is the header row.
        var headerIndex = rawRows.FindIndex(r => r.Any(c => !c.IsEmpty));
        if (headerIndex < 0)
        {
            throw new InvalidDataException("The worksheet contains no data.");
        }

        var headerRow = rawRows[headerIndex];
        var columnCount = maxColumn;
        var dataStart = headerIndex + 1;

        // Determine which columns are "live" - a column is kept only if it has a
        // header value or at least one non-empty data cell. This drops phantom
        // trailing columns that exist only because of cell formatting/styling.
        var isLive = new bool[columnCount];
        for (var i = 0; i < columnCount; i++)
        {
            var hasHeader = i < headerRow.Length && !headerRow[i].IsEmpty;
            if (hasHeader)
            {
                isLive[i] = true;
                continue;
            }

            for (var r = dataStart; r < rawRows.Count; r++)
            {
                var row = rawRows[r];
                if (i < row.Length && !row[i].IsEmpty)
                {
                    isLive[i] = true;
                    break;
                }
            }
        }

        var liveColumns = Enumerable.Range(0, columnCount).Where(i => isLive[i]).ToArray();
        if (liveColumns.Length == 0)
        {
            throw new InvalidDataException("The worksheet contains no usable columns.");
        }

        var headers = new List<string>(liveColumns.Length);
        for (var c = 0; c < liveColumns.Length; c++)
        {
            var i = liveColumns[c];
            var raw = i < headerRow.Length ? headerRow[i].AsText().Trim() : string.Empty;
            headers.Add(string.IsNullOrEmpty(raw) ? $"Column{c + 1}" : raw);
        }

        var dataRows = new List<ExcelCell[]>();
        for (var r = dataStart; r < rawRows.Count; r++)
        {
            var row = rawRows[r];
            if (row.All(c => c.IsEmpty))
            {
                continue; // skip fully blank rows
            }

            // Keep only the live columns, in their original order.
            var projected = new ExcelCell[liveColumns.Length];
            for (var c = 0; c < liveColumns.Length; c++)
            {
                var i = liveColumns[c];
                projected[c] = i < row.Length ? row[i] : ExcelCell.Empty;
            }

            dataRows.Add(projected);
        }

        return new ExcelSheet { Headers = headers, Rows = dataRows };
    }

    private static Sheet ResolveSheet(WorkbookPart workbookPart, string? worksheetName)
    {
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList()
            ?? new List<Sheet>();

        if (sheets.Count == 0)
        {
            throw new InvalidDataException("The workbook does not contain any worksheets.");
        }

        if (string.IsNullOrWhiteSpace(worksheetName))
        {
            return sheets[0];
        }

        var match = sheets.FirstOrDefault(s =>
            string.Equals(s.Name?.Value, worksheetName, StringComparison.OrdinalIgnoreCase));

        if (match is null)
        {
            var available = string.Join(", ", sheets.Select(s => $"'{s.Name?.Value}'"));
            throw new InvalidDataException(
                $"Worksheet '{worksheetName}' was not found. Available worksheets: {available}.");
        }

        return match;
    }

    private static string?[] LoadSharedStrings(WorkbookPart workbookPart)
    {
        var table = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (table is null)
        {
            return Array.Empty<string?>();
        }

        return table.Elements<SharedStringItem>()
            .Select(item => item.InnerText)
            .ToArray();
    }

    /// <summary>
    /// Returns the set of style indexes whose number format renders a date/time,
    /// so numeric cells using those styles can be surfaced as DateTime values.
    /// </summary>
    private static HashSet<uint> LoadDateFormatStyles(WorkbookPart workbookPart)
    {
        var result = new HashSet<uint>();
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        var cellFormats = stylesheet?.CellFormats;
        if (cellFormats is null)
        {
            return result;
        }

        var customFormats = stylesheet!.NumberingFormats?
            .Elements<NumberingFormat>()
            .ToDictionary(
                n => n.NumberFormatId?.Value ?? 0u,
                n => n.FormatCode?.Value ?? string.Empty)
            ?? new Dictionary<uint, string>();

        uint styleIndex = 0;
        foreach (var format in cellFormats.Elements<CellFormat>())
        {
            var numberFormatId = format.NumberFormatId?.Value ?? 0u;
            if (IsDateFormatId(numberFormatId) ||
                (customFormats.TryGetValue(numberFormatId, out var code) && LooksLikeDateFormat(code)))
            {
                result.Add(styleIndex);
            }

            styleIndex++;
        }

        return result;
    }

    private static bool IsDateFormatId(uint id) =>
        id is (>= 14 and <= 22) or (>= 45 and <= 47);

    private static bool LooksLikeDateFormat(string code)
    {
        if (string.IsNullOrEmpty(code))
        {
            return false;
        }

        // Strip quoted literals and colour/condition tokens before sniffing.
        var stripped = code.Replace("\"", string.Empty);
        return stripped.IndexOfAny(new[] { 'y', 'd', 'h', 's' }) >= 0
            || stripped.Contains('m'); // month/minute - close enough for inference
    }

    private static ExcelCell ParseCell(Cell cell, string?[] sharedStrings, HashSet<uint> dateStyles)
    {
        var rawValue = cell.CellValue?.InnerText;

        // Inline strings carry their text directly on the cell.
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            var text = cell.InlineString?.Text?.Text ?? cell.InnerText;
            return string.IsNullOrEmpty(text) ? ExcelCell.Empty : new ExcelCell(text, ExcelCellKind.Text);
        }

        if (string.IsNullOrEmpty(rawValue))
        {
            return ExcelCell.Empty;
        }

        // CellValues is a struct in the OpenXML SDK (3.x), so its members are not
        // compile-time constants and cannot be used in a switch; compare with ==.
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.SharedString)
        {
            if (int.TryParse(rawValue, out var ssIndex) && ssIndex >= 0 && ssIndex < sharedStrings.Length)
            {
                var text = sharedStrings[ssIndex];
                return string.IsNullOrEmpty(text) ? ExcelCell.Empty : new ExcelCell(text, ExcelCellKind.Text);
            }

            return ExcelCell.Empty;
        }

        if (dataType == CellValues.Boolean)
        {
            return new ExcelCell(rawValue == "1", ExcelCellKind.Boolean);
        }

        if (dataType == CellValues.Date)
        {
            return DateTime.TryParse(rawValue, out var isoDate)
                ? new ExcelCell(isoDate, ExcelCellKind.Date)
                : new ExcelCell(rawValue, ExcelCellKind.Text);
        }

        if (dataType == CellValues.String) // formula result string
        {
            return new ExcelCell(rawValue, ExcelCellKind.Text);
        }

        // Number (possibly a serial date depending on the cell style).
        if (double.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var number))
        {
            var styleIndex = cell.StyleIndex?.Value;
            if (styleIndex.HasValue && dateStyles.Contains(styleIndex.Value))
            {
                return new ExcelCell(DateTime.FromOADate(number), ExcelCellKind.Date);
            }

            return new ExcelCell(number, ExcelCellKind.Number);
        }

        return new ExcelCell(rawValue, ExcelCellKind.Text);
    }

    /// <summary>Converts a cell reference such as "B7" into a zero-based column index.</summary>
    private static int ColumnIndexFromReference(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
        {
            return -1;
        }

        var index = 0;
        var found = false;
        foreach (var ch in reference)
        {
            if (ch is >= 'A' and <= 'Z')
            {
                index = (index * 26) + (ch - 'A' + 1);
                found = true;
            }
            else if (ch is >= 'a' and <= 'z')
            {
                index = (index * 26) + (ch - 'a' + 1);
                found = true;
            }
            else
            {
                break;
            }
        }

        return found ? index - 1 : -1;
    }
}
