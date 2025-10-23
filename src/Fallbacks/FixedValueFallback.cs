using System.Reflection;

namespace ExcelMapper.Fallbacks;

/// <summary>
/// A fallback that returns a fixed given value when falling back.
/// </summary>
/// <param name="value">The fixed value returned when falling back.</param>
public class FixedValueFallback(object? value) : IFallbackItem
{
    /// <summary>
    /// The fixed value returned when falling back.
    /// </summary>
    public object? Value { get; } = value;

    /// <inheritdoc/>
    public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult result, Exception? exception, MemberInfo? member)
        => Value;
}
