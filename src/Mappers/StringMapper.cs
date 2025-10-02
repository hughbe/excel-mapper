using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that returns the string value of a cell.
/// </summary>
public class StringMapper : ICellMapper
{
    public CellMapperResult MapCellValue(ReadCellResult result)
        => CellMapperResult.SuccessIfNoOtherSuccess(result.StringValue);
}
