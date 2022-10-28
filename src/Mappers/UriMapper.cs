using System.Reflection;
using ExcelMapper.Abstractions;

/// <summary>
/// Tries to map the value of a cell to an absolute Uri.
/// </summary>
public class UriMapper : ICellValueMapper
{
    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        try
        {
            var uri = new Uri(previous.StringValue, UriKind.Absolute);
            return previous.Success(uri);
        }
        catch (Exception exception)
        {
            return previous.Invalid(exception);
        }
    }
}
