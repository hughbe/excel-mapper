using System.Reflection;

namespace ExcelMapper.Fallbacks;

/// <summary>
/// A fallback that throws an ExcelMappingException when falling back.
/// </summary>
public class ThrowFallback : IFallbackItem
{
    public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
    {
        var stringValue = readResult.GetString();
        if (member == null)
        {
            throw new ExcelMappingException($"Cannot assign \"{stringValue}\"", sheet, rowIndex, readResult.ColumnIndex, exception);
        }

        throw new ExcelMappingException($"Cannot assign \"{stringValue}\" to member \"{member.Name}\" of type \"{member.MemberType()}\"", sheet, rowIndex, readResult.ColumnIndex, exception);
    }
}
