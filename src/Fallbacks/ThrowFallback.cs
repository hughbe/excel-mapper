using System;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Fallbacks
{
    /// <summary>
    /// A fallback that throws an ExcelMappingException when falling back.
    /// </summary>
    public class ThrowFallback : IFallbackItem
    {
        public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellValueResult readResult, Exception exception, MemberInfo member)
        {
            throw new ExcelMappingException($"Invalid assigning \"{readResult.StringValue}\" to member \"{member.Name}\" of type \"{member.MemberType()}\"", sheet, rowIndex, readResult.ColumnIndex, exception);
        }
    }
}
