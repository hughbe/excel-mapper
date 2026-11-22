namespace ExcelMapper.Fallbacks;

/// <summary>
/// Uses an <see cref="ICellFallback{T}"/> for members of type <c>T?</c>.
/// </summary>
internal sealed class NullableCellFallback<T>(ICellFallback<T> fallback) : ICellFallback<T?> where T : struct
{
    public ICellFallback<T> Fallback { get; } = fallback;

    T? ICellFallback<T?>.Fallback(CellValue value, CellContext context, Exception? exception) => Fallback.Fallback(value, context, exception);
}
