namespace ExcelMapper.Fallbacks;

/// <summary>
/// A fallback that gives a fixed value, such as the value of <see cref="ExcelDefaultValueAttribute"/>
/// or <c>WithEmptyFallback(value)</c>.
/// </summary>
/// <remarks>
/// The value is given as an object (attributes and the untyped fluent methods can't know the
/// member's type), so a value that isn't a <typeparamref name="T"/> is only an error when it's
/// used: for example <c>null</c> for an <c>int</c> member throws an
/// <see cref="ExcelMappingException"/> for the cell.
/// </remarks>
internal sealed class FixedValueFallback<T> : ICellFallback<T>
{
    // The value as a T, if it is one, so that using it doesn't unbox it.
    private readonly T _typedValue = default!;
    private readonly bool _isTyped;

    public FixedValueFallback(object? value)
    {
        Value = value;
        if (value is T typedValue)
        {
            _typedValue = typedValue;
            _isTyped = true;
        }
        else if (value is null && default(T) is null)
        {
            _isTyped = true;
        }
    }

    public object? Value { get; }

    public T Fallback(CellValue value, CellContext context, Exception? exception)
        => _isTyped ? _typedValue : ValuePipeline.CastValue<T>(Value, context.ReadingSheet, context.RowIndex, context.ColumnIndex, context.Member);
}
