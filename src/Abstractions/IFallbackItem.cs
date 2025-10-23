using System.Reflection;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Defines a fallback item for handling read errors or missing data.
/// </summary>
public interface IFallbackItem
{
    /// <summary>
    /// Performs the fallback operation.
    /// </summary>
    /// <param name="sheet">The sheet.</param>
    /// <param name="rowIndex">The row index.</param>
    /// <param name="readResult">The read cell result.</param>
    /// <param name="exception">The exception that occurred, if any.</param>
    /// <param name="member">The member being mapped to, if any.</param>
    /// <returns>The fallback value.</returns>
    object? PerformFallback(
        ExcelSheet sheet,
        int rowIndex,
        ReadCellResult readResult,
        Exception? exception,
        MemberInfo? member);
}
