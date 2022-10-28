using System.Reflection;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Converts a cell to a value.
/// </summary> 
public interface ICellValueMapper
{
    CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member);
}
