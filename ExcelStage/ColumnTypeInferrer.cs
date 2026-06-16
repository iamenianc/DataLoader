using System.Data;
using System.Globalization;

namespace ExcelStage;

/// <summary>Describes a single inferred staging column.</summary>
public sealed class InferredColumn
{
    public required string Name { get; init; }
    public required string SqlType { get; init; }
    public required SqlDbType DbType { get; init; }
    public bool IsNullable { get; init; } = true;

    /// <summary>Converts a raw Excel cell into the CLR value SqlBulkCopy should send.</summary>
    public object? Convert(ExcelCell cell)
    {
        if (cell.IsEmpty)
        {
            return DBNull.Value;
        }

        try
        {
            return DbType switch
            {
                SqlDbType.Bit => ToBool(cell),
                SqlDbType.BigInt => System.Convert.ToInt64(ToNumber(cell)),
                SqlDbType.Decimal => ToDecimal(cell),
                SqlDbType.Float => System.Convert.ToDouble(ToNumber(cell)),
                SqlDbType.DateTime2 => ToDateTime(cell),
                _ => cell.AsText()
            };
        }
        catch
        {
            // Fall back to text so a single odd value never aborts the load.
            return cell.AsText();
        }
    }

    private static object ToBool(ExcelCell cell)
    {
        if (cell.Value is bool b) return b;
        var text = cell.AsText().Trim();
        if (bool.TryParse(text, out var parsed)) return parsed;
        return text is "1" or "yes" or "y" or "true";
    }

    private static object ToNumber(ExcelCell cell) =>
        cell.Value is double or long or int or decimal
            ? cell.Value!
            : double.Parse(cell.AsText(), NumberStyles.Any, CultureInfo.InvariantCulture);

    private static object ToDecimal(ExcelCell cell) =>
        cell.Value is double d
            ? (decimal)d
            : decimal.Parse(cell.AsText(), NumberStyles.Any, CultureInfo.InvariantCulture);

    private static object ToDateTime(ExcelCell cell) =>
        cell.Value is DateTime dt
            ? dt
            : DateTime.Parse(cell.AsText(), CultureInfo.InvariantCulture);
}

public static class ColumnTypeInferrer
{
    private const int MaxVarcharLength = 4000;

    public static List<InferredColumn> Infer(ExcelSheet sheet)
    {
        var columns = new List<InferredColumn>(sheet.Headers.Count);
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var col = 0; col < sheet.Headers.Count; col++)
        {
            var name = MakeUniqueName(sheet.Headers[col], usedNames);

            var allBool = true;
            var allInt = true;
            var allDecimal = true;
            var allDate = true;
            var hasValue = false;
            var maxLength = 1;

            foreach (var row in sheet.Rows)
            {
                var cell = col < row.Length ? row[col] : ExcelCell.Empty;
                if (cell.IsEmpty)
                {
                    continue;
                }

                hasValue = true;
                var text = cell.AsText();
                if (text.Length > maxLength)
                {
                    maxLength = text.Length;
                }

                allBool &= IsBool(cell);
                allInt &= IsInteger(cell);
                allDecimal &= IsDecimal(cell);
                allDate &= IsDate(cell);
            }

            columns.Add(BuildColumn(name, hasValue, allBool, allInt, allDecimal, allDate, maxLength));
        }

        return columns;
    }

    private static InferredColumn BuildColumn(
        string name, bool hasValue, bool allBool, bool allInt, bool allDecimal, bool allDate, int maxLength)
    {
        if (!hasValue)
        {
            return new InferredColumn { Name = name, SqlType = "NVARCHAR(255)", DbType = SqlDbType.NVarChar };
        }

        if (allBool)
        {
            return new InferredColumn { Name = name, SqlType = "BIT", DbType = SqlDbType.Bit };
        }

        if (allInt)
        {
            return new InferredColumn { Name = name, SqlType = "BIGINT", DbType = SqlDbType.BigInt };
        }

        if (allDecimal)
        {
            return new InferredColumn { Name = name, SqlType = "DECIMAL(38, 10)", DbType = SqlDbType.Decimal };
        }

        if (allDate)
        {
            return new InferredColumn { Name = name, SqlType = "DATETIME2", DbType = SqlDbType.DateTime2 };
        }

        if (maxLength > MaxVarcharLength)
        {
            return new InferredColumn { Name = name, SqlType = "NVARCHAR(MAX)", DbType = SqlDbType.NVarChar };
        }

        // Pad the inferred length to give a little headroom.
        var declared = Math.Min(MaxVarcharLength, Math.Max(50, maxLength + (maxLength / 5)));
        return new InferredColumn
        {
            Name = name,
            SqlType = $"NVARCHAR({declared})",
            DbType = SqlDbType.NVarChar
        };
    }

    private static bool IsBool(ExcelCell cell)
    {
        if (cell.Kind == ExcelCellKind.Boolean) return true;
        var text = cell.AsText().Trim().ToLowerInvariant();
        return text is "true" or "false" or "0" or "1" or "yes" or "no";
    }

    private static bool IsInteger(ExcelCell cell)
    {
        if (cell.Value is double d)
        {
            return Math.Abs(d % 1) <= double.Epsilon && d is >= long.MinValue and <= long.MaxValue;
        }

        return long.TryParse(cell.AsText(), NumberStyles.Integer, CultureInfo.InvariantCulture, out _);
    }

    private static bool IsDecimal(ExcelCell cell)
    {
        if (cell.Kind == ExcelCellKind.Number) return true;
        return decimal.TryParse(cell.AsText(), NumberStyles.Any, CultureInfo.InvariantCulture, out _);
    }

    private static bool IsDate(ExcelCell cell)
    {
        if (cell.Kind == ExcelCellKind.Date) return true;
        if (cell.Kind == ExcelCellKind.Number) return false; // a bare number is not a date
        return DateTime.TryParse(cell.AsText(), CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
    }

    private static string MakeUniqueName(string proposed, HashSet<string> used)
    {
        var name = SanitizeIdentifier(proposed);
        if (used.Add(name))
        {
            return name;
        }

        var suffix = 2;
        string candidate;
        do
        {
            candidate = $"{name}_{suffix++}";
        }
        while (!used.Add(candidate));

        return candidate;
    }

    private static string SanitizeIdentifier(string raw)
    {
        var trimmed = raw.Trim();
        if (string.IsNullOrEmpty(trimmed))
        {
            return "Column";
        }

        var chars = trimmed.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray();
        var name = new string(chars);
        if (char.IsDigit(name[0]))
        {
            name = "_" + name;
        }

        return name.Length > 128 ? name[..128] : name;
    }
}
