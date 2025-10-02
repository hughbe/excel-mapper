using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

public delegate CellMapperResult ConvertUsingMapperDelegate(ReadCellResult readResult);

/// <summary>
/// A mapper that tries to map the value of a cell to an object using a given conversion delegate.
/// </summary>
public class ConvertUsingMapper : ICellMapper
{
    /// <summary>
    /// Gets the delegate used to map the value of a cell to an object.
    /// </summary>
    public ConvertUsingMapperDelegate Converter { get; }

    /// <summary>
    /// Constructs a mapper that tries to map the value of a cell to an object using a given conversion delegate.
    /// </summary>
    /// <param name="converter">The delegate used to map the value of a cell to an object</param>
    public ConvertUsingMapper(ConvertUsingMapperDelegate converter)
    {
        Converter = converter ?? throw new ArgumentNullException(nameof(converter));
    }

    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        try
        {
            return Converter(readResult);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}
