namespace ExcelMapper.Mappers;

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
