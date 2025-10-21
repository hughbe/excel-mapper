using System;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to an IParsable object using IParsable.Parse.
/// </summary>
public class ParsableMapper<TParse> : ICellMapper where TParse : IParsable<TParse>
{
    /// <summary>
    /// Gets or sets the IFormatProvider used to map the value to the type.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        var value = readResult.GetString();
        try
        {
            var result = TParse.Parse(value!, Provider);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}
