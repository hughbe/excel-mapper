namespace ExcelMapper.Mappers;

/// <summary>
/// A delegate that maps the result of reading an Excel cell to a <see cref="CellMapperResult"/>.
/// </summary>
/// <param name="readResult">The result of reading a cell, containing the value and column information.</param>
/// <returns>A <see cref="CellMapperResult"/> indicating whether the mapping succeeded and the mapped value.</returns>
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
        ThrowHelpers.ThrowIfNull(converter, nameof(converter));
        Converter = converter;
    }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
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
