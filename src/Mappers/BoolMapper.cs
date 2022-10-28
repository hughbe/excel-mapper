using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;
    
/// <summary>
/// A mapper that tries to map the value of a cell to a bool.
/// </summary>
public class BoolMapper : ICellValueMapper
{
    private static object s_boxedTrue = true;
    private static object s_boxedFalse = false;

    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        if (previous.Value is bool boolValue)
        {
            return previous.Success(boolValue);
        }

        // Excel transforms bool values such as "true" or "false" to "1" or "0".
        string stringValue = previous.StringValue;
        if (stringValue == "1")
        {
            return previous.Success(s_boxedTrue);
        }
        else if (stringValue == "0")
        {
            return previous.Success(s_boxedFalse);
        }

        try
        {
            bool result = bool.Parse(stringValue);
            return previous.Success(result);
        }
        catch (Exception exception)
        {
            return previous.Invalid(exception);
        }
    }
}
