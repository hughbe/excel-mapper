using System.Reflection;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that needs to know where the cell is: <see cref="ValuePipeline.GetPropertyValue"/> calls
/// <see cref="Map(ReadCellResult, ExcelSheet, int, MemberInfo?)"/> rather than <see cref="ICellMapper.Map"/>.
/// </summary>
internal interface IContextCellMapper : ICellMapper
{
    CellMapperResult Map(ReadCellResult readResult, ExcelSheet sheet, int rowIndex, MemberInfo? member);
}

/// <summary>
/// The mapper for an <see cref="ICellConverter{T}"/> set with <c>WithConverter(converter)</c> or
/// <see cref="ExcelConverterAttribute"/>. A <see cref="ValuePipeline{T}"/> calls the converter
/// directly, without boxing (see <see cref="Converter"/>); this mapper is used where values go
/// through the pipeline as objects.
/// </summary>
internal sealed class CellConverterMapper<T>(ICellConverter<T> converter) : IContextCellMapper
{
    public ICellConverter<T> Converter { get; } = converter;

    // Without a sheet, for example if the mapper is called directly.
    public CellMapperResult Map(ReadCellResult readResult) => ToResult(Converter.Convert(new CellValue(readResult), default));

    public CellMapperResult Map(ReadCellResult readResult, ExcelSheet sheet, int rowIndex, MemberInfo? member)
        => ToResult(Converter.Convert(new CellValue(readResult), new CellContext(sheet, rowIndex, readResult.ColumnIndex, member)));

    private static CellMapperResult ToResult(CellConversionResult<T> result) => result.Status switch
    {
        CellConversionStatus.Success => CellMapperResult.Success(result.Value),
        CellConversionStatus.Invalid => result.Error switch
        {
            Func<Exception> createException => CellMapperResult.Invalid(createException),
            var error => new CellMapperResult(null, (Exception?)error, CellMapperResult.HandleAction.ErrorAndContinueMapping),
        },
        _ => CellMapperResult.Ignore(),
    };
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
internal sealed class DelegateCellConverter<T>(ConvertUsingSimpleMapperDelegate<T> converter) : ICellConverter<T>
{
    public ConvertUsingSimpleMapperDelegate<T> Converter { get; } = converter;

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
