namespace ExcelStage;

/// <summary>The kind of result a wizard step produced.</summary>
public enum NavKind
{
    Value,    // the user supplied a value
    Back,     // go to the previous step
    Restart,  // start over from the first step
    Quit      // cancel the whole run
}

/// <summary>
/// The outcome of a single wizard step: either a value, or a navigation
/// request (back / restart / quit).
/// </summary>
public readonly struct Nav<T>
{
    private Nav(NavKind kind, T? value)
    {
        Kind = kind;
        Value = value;
    }

    public NavKind Kind { get; }
    public T? Value { get; }

    public bool IsValue => Kind == NavKind.Value;

    public static Nav<T> FromValue(T value) => new(NavKind.Value, value);
    public static Nav<T> Back { get; } = new(NavKind.Back, default);
    public static Nav<T> Restart { get; } = new(NavKind.Restart, default);
    public static Nav<T> Quit { get; } = new(NavKind.Quit, default);

    /// <summary>Re-wraps a non-value result as a different element type.</summary>
    public Nav<TOther> Carry<TOther>() => Kind switch
    {
        NavKind.Back => Nav<TOther>.Back,
        NavKind.Restart => Nav<TOther>.Restart,
        _ => Nav<TOther>.Quit
    };
}
