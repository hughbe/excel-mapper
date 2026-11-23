namespace ExcelMapper.Fallbacks;

/// <summary>
/// The default invalid fallback, and the empty fallback of non-nullable value types: reports the
/// cell as an error.
/// </summary>
/// <remarks>
/// When errors are collected rather than thrown, <see cref="CellPipeline{T}.Map"/> reports the
/// cell to them itself instead of calling this, so that the exception is only created if it's
/// needed: throwing and catching an exception for every invalid cell would be slow.
/// </remarks>
internal sealed class ThrowFallback<T> : ICellFallback<T>
{
    public static ThrowFallback<T> Instance { get; } = new();

    /// <summary>
    /// What to do about an empty cell, made once for the type.
    /// </summary>
    public static string EmptyHint { get; } = CellErrors.GetEmptyHint(typeof(T));

    public T Fallback(CellValue value, CellContext context, Exception? exception)
        => throw CellErrors.Create(value, context, exception, typeof(T), EmptyHint);
}
