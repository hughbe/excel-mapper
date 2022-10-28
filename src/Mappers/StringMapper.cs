using System.Reflection;
using ExcelMapper.Abstractions;

/// <summary>
/// A mapper that returns the string value of a cell.
/// </summary>
public class StringMapper : ICellValueMapper
{
    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        return CellValueMapperResult.SuccessIfNoOtherSuccess(previous.Value?.ToString());
    }
}
