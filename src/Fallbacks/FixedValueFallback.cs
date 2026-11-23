namespace ExcelMapper.Fallbacks;

/// <summary>
/// A fallback that gives a fixed value, such as the value of <see cref="ExcelDefaultValueAttribute"/>
/// or <c>WithEmptyFallback(value)</c>.
/// </summary>
/// <remarks>
/// The value is a <typeparamref name="T"/>: attribute values are converted to the member's type
/// when the class map is created (see <see cref="AutoMapping.MemberAttributes.ConvertValue{T}"/>).
/// </remarks>
internal sealed class FixedValueFallback<T>(T value) : ICellFallback<T>
{
    public T Value { get; } = value;

    public T Fallback(CellValue value, CellContext context, Exception? exception) => Value;
}
