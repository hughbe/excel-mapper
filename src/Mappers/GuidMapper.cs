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
        try
        {
            // Discarding readResult.StringValue nullability warning.
            // If null - CellMapperResult.Invalid with ArgumentNullException will be returned
            Guid result = Guid.Parse(readResult.StringValue!);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}
