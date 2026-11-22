namespace ExcelMapper.Mappers;

/// <summary>
/// The mapper for an <see cref="ICellConverter{T}"/> set with <c>WithConverter(converter)</c> or
/// <see cref="ExcelConverterAttribute"/>. The error the converter returns for an invalid cell is
/// kept for <see cref="CreateException"/>, as the converter can't be asked for it again without
/// converting the value again.
/// </summary>
internal sealed class CellConverterMapper<T>(ICellConverter<T> converter) : ICellMapper<T>
{
    public ICellConverter<T> Converter { get; } = converter;

    public CellConversion Convert(ReadCellResult cell, in CellContext context, out T value, ref object? error)
    {
        var result = Converter.Convert(new CellValue(cell), context);
        value = result.Value;
        switch (result.Status)
        {
            case CellConversionStatus.Success:
                return CellConversion.Success;
            case CellConversionStatus.Invalid:
                error = result.Error;
                return CellConversion.Invalid;
            default:
                return CellConversion.NotHandled;
        }
    }

    public Exception? CreateException(ReadCellResult cell, object? error) => CellConversionResult<T>.CreateException(error);
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
