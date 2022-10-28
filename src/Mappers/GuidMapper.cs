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
            return previous.Success(guidValue);
        }

        try
        {
            Guid result = Guid.Parse(previous.StringValue);
            return previous.Success(result);
        }
        catch (Exception exception)
        {
            return previous.Invalid(exception);
        }
    }
}
