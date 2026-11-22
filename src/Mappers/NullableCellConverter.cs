namespace ExcelMapper.Mappers;

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
