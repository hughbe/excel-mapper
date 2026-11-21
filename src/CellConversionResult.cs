namespace ExcelMapper;

/// <summary>
/// Whether an <see cref="ICellConverter{T}"/> converted a cell. See <see cref="CellConversionResult{T}"/>.
/// </summary>
public enum CellConversionStatus
{
    /// <summary>
    /// The converter doesn't handle the value, so the next converter is tried. If no converter
    /// handles it, the value is invalid. This is the status of <c>default(CellConversionResult&lt;T&gt;)</c>.
    /// </summary>
    NotHandled,

    /// <summary>The value was converted.</summary>
    Success,

    /// <summary>
    /// The value is invalid. The next converter is still tried; if none converts the value, the
    /// member's invalid fallback is used, which by default throws an <see cref="ExcelMappingException"/>.
    /// </summary>
    Invalid,
}

/// <summary>
/// The result of converting a cell with an <see cref="ICellConverter{T}"/>.
/// </summary>
/// <typeparam name="T">The type the cell is converted to.</typeparam>
public readonly struct CellConversionResult<T>
{
    // An Exception, or a Func<Exception> that creates it.
    private readonly object? _error;

    private CellConversionResult(CellConversionStatus status, T value, object? error)
    {
        Status = status;
        Value = value;
        _error = error;
    }

    /// <summary>
    /// Gets whether the cell was converted.
    /// </summary>
    public CellConversionStatus Status { get; }

    /// <summary>
    /// Gets the converted value, if <see cref="Status"/> is <see cref="CellConversionStatus.Success"/>.
    /// </summary>
    public T Value { get; }

    /// <summary>
    /// Gets the exception that describes why the value is invalid, if there is one. An exception
    /// given to <see cref="Invalid(Func{Exception})"/> is created each time this is read.
    /// </summary>
    public Exception? Exception => CreateException(_error);

    /// <summary>
    /// The cell was converted to <paramref name="value"/>.
    /// </summary>
    public static CellConversionResult<T> Success(T value) => new(CellConversionStatus.Success, value, null);

    /// <summary>
    /// The value is invalid, without an exception to describe why.
    /// </summary>
    public static CellConversionResult<T> Invalid() => new(CellConversionStatus.Invalid, default!, null);

    /// <summary>
    /// The value is invalid, for the reason <paramref name="exception"/> gives. It's the inner
    /// exception of the <see cref="ExcelMappingException"/> thrown for the cell.
    /// </summary>
    public static CellConversionResult<T> Invalid(Exception exception)
    {
        ThrowHelpers.ThrowIfNull(exception, nameof(exception));
        return new(CellConversionStatus.Invalid, default!, exception);
    }

    /// <summary>
    /// The value is invalid, for the reason the exception <paramref name="createException"/>
    /// creates. The exception is only created if it's used, for example by the default invalid
    /// fallback, which throws. Use this when invalid values are common and are given a fallback
    /// value, so that creating exceptions doesn't slow mapping down.
    /// </summary>
    public static CellConversionResult<T> Invalid(Func<Exception> createException)
    {
        ThrowHelpers.ThrowIfNull(createException, nameof(createException));
        return new(CellConversionStatus.Invalid, default!, createException);
    }

    /// <summary>
    /// The converter doesn't handle the value, so the next converter is tried.
    /// </summary>
    public static CellConversionResult<T> NotHandled => default;

    // The exception or the function that creates it, without creating it.
    internal object? Error => _error;

    internal static CellConversionResult<T> InvalidWithError(object? error) => new(CellConversionStatus.Invalid, default!, error);

    // Creates the exception from an error kept by a result.
    internal static Exception? CreateException(object? error) => error switch
    {
        Exception exception => exception,
        Func<Exception> createException => createException(),
        _ => null,
    };
}
