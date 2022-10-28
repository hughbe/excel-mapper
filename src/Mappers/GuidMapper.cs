using System.Reflection;
using ExcelMapper.Abstractions;

/// <summary>
/// A mapper that tries to map the value of a cell to a guid.
/// </summary>
public class GuidMapper : ICellValueMapper
{
    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        if (previous.Value is Guid guidValue)
        {
            return CellValueMapperResult.Success(guidValue);
        }

        string stringValue = previous.Value?.ToString();
        try
        {
            Guid result = Guid.Parse(stringValue);
            return CellValueMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellValueMapperResult.Invalid(exception);
        }
    }
}
