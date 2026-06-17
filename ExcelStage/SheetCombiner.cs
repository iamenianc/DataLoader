namespace ExcelStage;

/// <summary>
/// Combines one or more worksheets (selected from the same workbook) into a
/// single logical sheet ready for staging.
///
/// Columns are unioned by name: a column that appears in more than one sheet is
/// added to the table only once. Each column's SQL type is inferred from the
/// FIRST selected sheet that contains it, and every selected sheet's data rows
/// are appended - cells are left empty (loaded as NULL) where a sheet does not
/// contain a given column.
/// </summary>
public static class SheetCombiner
{
    public sealed class Combined
    {
        /// <summary>The unioned columns, in first-appearance order.</summary>
        public required List<InferredColumn> Columns { get; init; }

        /// <summary>All rows from every sheet, projected onto <see cref="Columns"/>.</summary>
        public required ExcelSheet Sheet { get; init; }
    }

    /// <param name="sheets">
    /// The selected sheets, in selection order. The first entry drives column
    /// types for any column it contains.
    /// </param>
    public static Combined Combine(IReadOnlyList<ExcelSheet> sheets)
    {
        if (sheets.Count == 0)
        {
            throw new ArgumentException("At least one worksheet must be selected.", nameof(sheets));
        }

        // Infer each sheet on its own so a column's type comes from the first
        // selected sheet that contains it (not from the combined data).
        var perSheetColumns = sheets.Select(ColumnTypeInferrer.Infer).ToList();

        var columns = new List<InferredColumn>();
        var indexByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // sheetMaps[s][c] = the union column index that sheet s, column c maps to.
        var sheetMaps = new List<int[]>(sheets.Count);

        for (var s = 0; s < sheets.Count; s++)
        {
            var cols = perSheetColumns[s];
            var map = new int[cols.Count];
            for (var c = 0; c < cols.Count; c++)
            {
                var name = cols[c].Name;
                if (!indexByName.TryGetValue(name, out var unionIndex))
                {
                    unionIndex = columns.Count;
                    indexByName[name] = unionIndex;
                    columns.Add(cols[c]); // first sheet to define this column wins its type
                }

                map[c] = unionIndex;
            }

            sheetMaps.Add(map);
        }

        var width = columns.Count;
        var rows = new List<ExcelCell[]>();
        for (var s = 0; s < sheets.Count; s++)
        {
            var map = sheetMaps[s];
            foreach (var row in sheets[s].Rows)
            {
                var merged = new ExcelCell[width];
                Array.Fill(merged, ExcelCell.Empty);
                for (var c = 0; c < map.Length; c++)
                {
                    merged[map[c]] = c < row.Length ? row[c] : ExcelCell.Empty;
                }

                rows.Add(merged);
            }
        }

        var sheet = new ExcelSheet
        {
            Headers = columns.Select(col => col.Name).ToList(),
            Rows = rows
        };

        return new Combined { Columns = columns, Sheet = sheet };
    }
}
