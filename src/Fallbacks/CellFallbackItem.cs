using System.Reflection;

namespace ExcelMapper.Fallbacks;

/// <summary>
/// The pipeline fallback for an <see cref="ICellFallback{T}"/>. A <see cref="ValuePipeline{T}"/>
/// calls the fallback directly, without boxing (see <see cref="Fallback"/>).
/// </summary>
internal sealed class CellFallbackItem<T>(ICellFallback<T> fallback) : IFallbackItem
{
    public ICellFallback<T> Fallback { get; } = fallback;

    public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
        => Fallback.Fallback(new CellValue(readResult), new CellContext(sheet, rowIndex, readResult.ColumnIndex, member), exception);
}

/// <summary>
/// Uses an <see cref="ICellFallback{T}"/> for members of type <c>T?</c>.
/// </summary>
internal sealed class NullableCellFallback<T>(ICellFallback<T> fallback) : ICellFallback<T?> where T : struct
{
    public ICellFallback<T> Fallback { get; } = fallback;

    T? ICellFallback<T?>.Fallback(CellValue value, CellContext context, Exception? exception) => Fallback.Fallback(value, context, exception);
}
