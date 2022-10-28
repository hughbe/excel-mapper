using System.Reflection;

namespace ExcelMapper.Abstractions;

/// <summary>
/// An interface called when an invalid cell value is encountered. This can be used to return a fixed value or to throw.
/// </summary>
public interface IInvalidCellFallback
{
    object PerformFallback(ExcelCell cell, object value, Exception exception, MemberInfo member);
}
