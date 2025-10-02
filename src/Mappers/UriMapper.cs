using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to an absolute Uri.
/// </summary>
public class UriMapper : ICellMapper
{
    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        try
        {
            // Discarding readResult.StringValue nullability warning.
            // If null - CellMapperResult.Invalid with ArgumentNullException will be returned
            var uri = new Uri(readResult.StringValue!, UriKind.Absolute);
            return CellMapperResult.Success(uri);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}
