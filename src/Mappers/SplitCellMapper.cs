#if MULTI
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell and produces multiple values by splitting the string value
/// using the given separators.
/// </summary>
public abstract class SplitCellValueReader : ICellValueMapper
{
    /// <summary>
    /// Gets or sets the options used to split the string value of the cell.
    /// </summary>
    public StringSplitOptions Options { get; set; }

    public IEnumerable<CellValueMapperResult> MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        if (cell.Value == null)
        {
            return Enumerable.Empty<CellValueMapperResult>();
        }

        string stringValue = previous.Value?.ToString();
        return stringValue.Split(GetValues(stringValue), Options).Select(value => CellValueMapperResult.Success(value));
    }

    protected abstract string[] GetValues(string value);
}
#endif
