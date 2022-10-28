using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Fallbacks;
 
/// <summary>
/// A fallback that throws an ExcelMappingException when falling back.
/// </summary>
public class ThrowFallback : IEmptyCellFallback, IInvalidCellFallback
{
    object IEmptyCellFallback.PerformFallback(ExcelCell cell, MemberInfo member)
        => throw new ExcelMappingException($"Member \"{member.Name}\" of type \"{member.MemberType()}\" cannot be empty.", cell.Sheet, cell.RowIndex, cell.ColumnIndex);
    
    object IInvalidCellFallback.PerformFallback(ExcelCell cell, object value, Exception exception, MemberInfo member)
        => throw new ExcelMappingException($"Invalid assigning \"{value}\" to member \"{member.Name}\" of type \"{member.MemberType()}\"", cell.Sheet, cell.RowIndex, cell.ColumnIndex, exception);
}
