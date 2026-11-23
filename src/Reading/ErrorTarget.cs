namespace ExcelMapper.Reading;

/// <summary>
/// Where mapping errors go when they're collected instead of thrown: a list of exceptions for
/// <see cref="ExcelSheet.ReadRowResults{T}()"/>, or the user's <see cref="IExcelErrorSink"/>. Code
/// that can report a mapping error takes the target and adds to it when it isn't null.
/// </summary>
internal abstract class ErrorTarget
{
    /// <summary>
    /// Gets the number of errors added, so that code can tell whether reading a value added any.
    /// </summary>
    public int Count { get; private protected set; }

    /// <summary>
    /// Adds an error that already has its exception, such as one thrown by a constructor.
    /// </summary>
    public void Add(ExcelMappingException exception, ExcelMappingErrorKind kind = ExcelMappingErrorKind.Exception)
    {
        Count++;
        OnError(new ExcelMappingError(kind, exception));
    }

    /// <summary>
    /// Adds an empty or invalid cell. Its exception is only created if the target needs it.
    /// </summary>
    /// <param name="kind">Whether the cell is empty or invalid.</param>
    /// <param name="cell">The cell.</param>
    /// <param name="context">The cell's context.</param>
    /// <param name="valueType">The type the cell was converted to.</param>
    /// <param name="mapper">The mapper that found the cell invalid, which creates the reason.</param>
    /// <param name="mapperError">The error the mapper gave, if any.</param>
    /// <param name="emptyHint">What to do about an empty cell (see <see cref="CellErrors.GetEmptyHint"/>).</param>
    public virtual void AddCell(ExcelMappingErrorKind kind, ReadCellResult cell, in CellContext context, Type valueType, ICellExceptionSource? mapper, object? mapperError, string emptyHint)
    {
        Count++;
        OnError(new ExcelMappingError(kind, cell, context, valueType, mapper, mapperError, emptyHint));
    }

    private protected abstract void OnError(in ExcelMappingError error);
}

/// <summary>
/// Collects each error as an <see cref="ExcelMappingException"/>, for
/// <see cref="ExcelSheet.ReadRowResults{T}()"/>.
/// </summary>
internal sealed class ExceptionListTarget : ErrorTarget
{
    /// <summary>
    /// Gets the exceptions of the current row, which the reader copies and clears for each row.
    /// </summary>
    public List<ExcelMappingException> Exceptions { get; } = [];

    private protected override void OnError(in ExcelMappingError error) => Exceptions.Add(error.ToException());

    // Creates the exception directly, as every cell error needs one here.
    public override void AddCell(ExcelMappingErrorKind kind, ReadCellResult cell, in CellContext context, Type valueType, ICellExceptionSource? mapper, object? mapperError, string emptyHint)
    {
        Count++;
        var inner = kind == ExcelMappingErrorKind.InvalidValue ? mapper?.CreateException(cell, in context, mapperError) : null;
        Exceptions.Add(CellErrors.Create(new CellValue(cell), in context, inner, valueType, emptyHint));
    }
}

/// <summary>
/// Passes each error to the user's sink as it's found, without creating an exception.
/// </summary>
internal sealed class SinkTarget(IExcelErrorSink sink) : ErrorTarget
{
    private protected override void OnError(in ExcelMappingError error) => sink.Add(in error);
}
