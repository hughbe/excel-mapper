using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Fallbacks;
 
/// <summary>
/// A fallback that returns a fixed given value when falling back.
/// </summary>
public class FixedValueFallback : IEmptyCellFallback, IInvalidCellFallback
{
    /// <summary>
    /// The fixed value returned when falling back.
    /// </summary>
    public object Value { get; }

    /// <summary>
    /// Constructs a fallback that returns a given value when falling back.
    /// </summary>
    /// <param name="value">The fixed value returned when falling back.</param>
    public FixedValueFallback(object value) => Value = value;

    object IEmptyCellFallback.PerformFallback(ExcelCell cell, MemberInfo member) => Value;
    
    object IInvalidCellFallback.PerformFallback(ExcelCell cell, object value, Exception exception, MemberInfo member) => Value;
}
