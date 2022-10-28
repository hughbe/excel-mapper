using System.Reflection;

namespace ExcelMapper.Abstractions;

/// <summary>
/// An interface called when an empty cell is encountered. This can be used to return a fixed value or to throw.
/// </summary>
public interface IEmptyCellFallback
{
    object PerformFallback(ExcelCell cell, MemberInfo member);
}
