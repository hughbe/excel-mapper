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

    public Exception? CreateException(ReadCellResult cell, in CellContext context, object? error) => CellConversionResult<T>.CreateException(error);
}
