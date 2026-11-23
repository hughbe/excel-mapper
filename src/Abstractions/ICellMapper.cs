namespace ExcelMapper.Abstractions;

/// <summary>
/// The result of <see cref="ICellMapper{T}.Convert"/>.
/// </summary>
internal enum CellConversion
{
    /// <summary>
    /// The value was converted, and later mappers aren't tried.
    /// </summary>
    Success,

    /// <summary>
    /// The value was converted, but a later mapper that converts it replaces the value. Used for
    /// strings, so that a mapping dictionary after the string mapper wins.
    /// </summary>
    SuccessUnlessReplaced,

    /// <summary>
    /// The value is invalid. Later mappers are still tried.
    /// </summary>
    Invalid,

    /// <summary>
    /// The mapper doesn't handle the value, so later mappers are tried.
    /// </summary>
    NotHandled,
}

/// <summary>
/// Converts a cell to <typeparamref name="T"/> without boxing. A member's
/// <see cref="ValuePipeline{T}"/> tries its mappers in order.
/// </summary>
/// <typeparam name="T">The type of the member, or of the elements or values of a collection.</typeparam>
internal interface ICellMapper<T> : ICellExceptionSource
{
    /// <summary>
    /// Converts a cell's value.
    /// </summary>
    /// <param name="cell">The cell, after any transformers.</param>
    /// <param name="context">Where the cell is, for an <see cref="ICellConverter{T}"/>.</param>
    /// <param name="value">The converted value, if the result is a success.</param>
    /// <param name="error">For an invalid cell, an error to give <see cref="ICellExceptionSource.CreateException"/>, such
    /// as the one an <see cref="ICellConverter{T}"/> returned. Built-in mappers leave it.</param>
    /// <returns>Whether the value was converted.</returns>
    CellConversion Convert(ReadCellResult cell, in CellContext context, out T value, ref object? error);
}

/// <summary>
/// Creates the exception for a cell a mapper found invalid. It isn't generic, so that an error
/// reported to an <see cref="IExcelErrorSink"/> can create it later (see
/// <see cref="ExcelMappingError.ToException"/>).
/// </summary>
internal interface ICellExceptionSource
{
    /// <summary>
    /// Creates the exception that says why <see cref="ICellMapper{T}.Convert"/> found a cell invalid. It's only
    /// called when the exception is used (by the default invalid fallback, a custom fallback or
    /// <see cref="ExcelSheet.ReadRowResults{T}()"/>), so invalid cells with a fixed fallback don't
    /// create an exception at all.
    /// </summary>
    /// <param name="cell">The cell that <see cref="ICellMapper{T}.Convert"/> found invalid.</param>
    /// <param name="context">The cell's context, as given to <see cref="ICellMapper{T}.Convert"/>.</param>
    /// <param name="error">The error <see cref="ICellMapper{T}.Convert"/> gave, if any.</param>
    /// <returns>The exception, which isn't thrown, or null if there's none (a converter gave no error).</returns>
    Exception? CreateException(ReadCellResult cell, in CellContext context, object? error);
}
