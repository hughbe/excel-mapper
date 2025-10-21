using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to an absolute Uri.
/// </summary>
public class UriMapper : ICellMapper
{
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        try
            {
                // Discarding readResult.StringValue nullability warning.
                // If null - CellMapperResult.Invalid with ArgumentNullException will be returned
                var uri = new Uri(stringValue!, UriKind.Absolute);
                return CellMapperResult.Success(uri);
            }
            catch (Exception exception)
            {
                return CellMapperResult.Invalid(exception);
            }
    }
}
