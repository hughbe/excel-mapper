using System.Reflection;
using ExcelMapper.Abstractions;

/// <summary>
/// Tries to map the value of a cell to an absolute Uri.
/// </summary>
public class UriMapper : ICellValueMapper
{
    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        string stringValue = previous.Value?.ToString();
        try
        {
            var uri = new Uri(stringValue, UriKind.Absolute);
            return CellValueMapperResult.Success(uri);
        }
        catch (Exception exception)
        {
            return CellValueMapperResult.Invalid(exception);
        }
    }
}
