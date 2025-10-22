using System.Reflection;

namespace ExcelMapper.Fallbacks;

/// <summary>
/// A fallback that returns a fixed value produced by a factory function when falling back.
/// </summary>
public class FixedValueFallbackFactory : IFallbackItem
{
    /// <summary>
    /// The factory function that produces the fixed value.
    /// </summary>
    public Func<object?> Factory { get; }

    public FixedValueFallbackFactory(Func<object?> factory)
    {
        ArgumentNullException.ThrowIfNull(factory);
        Factory = factory;
    }

    public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult result, Exception? exception, MemberInfo? member)
        => Factory();
}
