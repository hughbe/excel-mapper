namespace ExcelMapper.Mappers;

/// <summary>
/// The mapper for an <see cref="ICellConverter{T}"/> set with <c>WithConverter(converter)</c> or
/// <see cref="ExcelConverterAttribute"/>. A <see cref="ValuePipeline{T}"/> calls the converter
/// itself, with the cell's context, and keeps the error it returns (see <see cref="Converter"/>);
/// this class only carries the converter in the pipeline's list of mappers.
/// </summary>
internal sealed class CellConverterMapper<T>(ICellConverter<T> converter) : ICellMapper<T>
{
    public ICellConverter<T> Converter { get; } = converter;

    // Not called by ValuePipeline<T>, which converts with the cell's context.
    CellConversion ICellMapper<T>.Convert(ReadCellResult cell, out T value)
        => throw new InvalidOperationException("Cell converters are called by the pipeline.");

    Exception ICellMapper<T>.CreateException(ReadCellResult cell)
        => throw new InvalidOperationException("Cell converters are called by the pipeline.");
}

/// <summary>
/// Uses an <see cref="ICellConverter{T}"/> for members of type <c>T?</c>.
/// </summary>
internal sealed class NullableCellConverter<T>(ICellConverter<T> converter) : ICellConverter<T?> where T : struct
{
    public ICellConverter<T> Converter { get; } = converter;

    public CellConversionResult<T?> Convert(CellValue value, CellContext context)
    {
        var result = Converter.Convert(value, context);
        return result.Status switch
        {
            CellConversionStatus.Success => CellConversionResult<T?>.Success(result.Value),
            CellConversionStatus.Invalid => CellConversionResult<T?>.InvalidWithError(result.Error),
            _ => CellConversionResult<T?>.NotHandled,
        };
    }
}

/// <summary>
/// The converter for <c>WithConverter(value => ...)</c>: converts a cell's text with a delegate. A
/// delegate that throws makes the value invalid, with the exception it threw.
/// </summary>
internal sealed class DelegateCellConverter<T>(Func<string?, T> converter) : ICellConverter<T>
{
    public Func<string?, T> Converter { get; } = converter;

    public CellConversionResult<T> Convert(CellValue value, CellContext context)
    {
        try
        {
            return CellConversionResult<T>.Success(Converter(value.Text));
        }
        catch (Exception exception)
        {
            return CellConversionResult<T>.Invalid(exception);
        }
    }
}
