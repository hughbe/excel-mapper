using System;
using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Fallbacks;

/// <summary>
/// A fallback that returns a fixed given value when falling back.
/// </summary>
/// <remarks>
/// Constructs a fallback that returns a given value when falling back.
/// </remarks>
/// <param name="value">The fixed value returned when falling back.</param>
public class FixedValueFallback(object? value) : IFallbackItem
{
    /// <summary>
    /// The fixed value returned when falling back.
    /// </summary>
    public object? Value { get; } = value;

    public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult result, Exception? exception, MemberInfo? member)
        => Value;
}
