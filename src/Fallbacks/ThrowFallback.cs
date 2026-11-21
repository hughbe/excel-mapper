using System.Reflection;

namespace ExcelMapper.Fallbacks;

/// <summary>
/// A fallback that throws an ExcelMappingException when falling back.
/// </summary>
public class ThrowFallback : IFallbackItem
{
    /// <inheritdoc/>
    public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
    {
        var stringValue = readResult.GetString();
        var mappingException = member == null
            ? new ExcelMappingException($"Cannot assign \"{stringValue}\"", sheet, rowIndex, readResult.ColumnIndex, exception)
            : new ExcelMappingException($"Cannot assign \"{stringValue}\" to member \"{member.Name}\" of type \"{member.MemberType()}\"", sheet, rowIndex, readResult.ColumnIndex, exception)
            {
                Member = member
            };

        // When the sheet is collecting errors (ReadRowResults), record the error rather than
        // throwing it: throwing and catching an exception for every invalid cell is slow.
        if (sheet?.CollectedErrors is { } errors)
        {
            errors.Add(mappingException);
            return ErrorRecorded;
        }

        throw mappingException;
    }

    /// <summary>
    /// Returned instead of throwing when an error has been recorded. Maps treat it as "no value".
    /// </summary>
    internal static readonly object ErrorRecorded = new();
}
