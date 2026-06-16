namespace ExcelStage;

/// <summary>
/// A single cell value read from the worksheet, carrying enough type
/// information for the column type inferrer to make a decision.
/// </summary>
public sealed class ExcelCell
{
    public static readonly ExcelCell Empty = new(null, ExcelCellKind.Empty);

    public ExcelCell(object? value, ExcelCellKind kind)
    {
        Value = value;
        Kind = kind;
    }

    /// <summary>The parsed value (string, double, bool, or DateTime), or null when empty.</summary>
    public object? Value { get; }

    public ExcelCellKind Kind { get; }

    public bool IsEmpty => Kind == ExcelCellKind.Empty || Value is null;

    public string AsText() => Value?.ToString() ?? string.Empty;
}

public enum ExcelCellKind
{
    Empty,
    Text,
    Number,
    Boolean,
    Date
}
