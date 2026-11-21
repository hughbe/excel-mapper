namespace ExcelMapper;

/// <summary>
/// Gives the value of a member whose cell is empty or invalid. Set it with
/// <c>WithEmptyFallback(fallback)</c>, <c>WithInvalidFallback(fallback)</c>,
/// <see cref="ExcelEmptyFallbackAttribute"/> or <see cref="ExcelInvalidFallbackAttribute"/>.
/// </summary>
/// <typeparam name="T">The type of the member. A fallback for a value type <c>T</c> can also be
/// used for members of type <c>T?</c>.</typeparam>
/// <remarks>
/// To report the cell as an error instead, throw an exception: an <see cref="ExcelMappingException"/>
/// is reported as is, and is recorded rather than thrown by <see cref="ExcelSheet.ReadRowResults{T}"/>.
/// A fallback can be used for many cells at once on different threads, so it mustn't change its own
/// state.
/// </remarks>
public interface ICellFallback<T>
{
    /// <summary>
    /// Gives the member's value for an empty or invalid cell.
    /// </summary>
    /// <param name="value">The cell's value, after any transformers.</param>
    /// <param name="context">Where the cell is, and the member it's mapped to.</param>
    /// <param name="exception">Why the value is invalid, if known. Null for an empty cell.</param>
    /// <returns>The member's value.</returns>
    T Fallback(CellValue value, CellContext context, Exception? exception);
}
