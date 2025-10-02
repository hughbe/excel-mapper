using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a guid.
/// </summary>
public class GuidMapper : ICellMapper
{
    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        try
        {
            // Discarding readResult.StringValue nullability warning.
            // If null - CellMapperResult.Invalid with ArgumentNullException will be returned
            Guid result = Guid.Parse(stringValue!);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}
