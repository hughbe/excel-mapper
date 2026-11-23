using System.Reflection;

namespace ExcelMapper;

/// <summary>
/// A problem found while reading a row, given to an <see cref="IExcelErrorSink"/>. It's only valid
/// during <see cref="IExcelErrorSink.Add"/>, as it refers to the cell while the reader is on its
/// row: copy what you need, or call <see cref="ToException"/>.
/// </summary>
public readonly ref struct ExcelMappingError
{
    private readonly ExcelMappingException? _exception;
    private readonly ReadCellResult _cell;
    private readonly CellContext _context;
    private readonly Type? _valueType;
    private readonly ICellExceptionSource? _mapper;
    private readonly object? _mapperError;
    private readonly string? _emptyHint;

    // An error that already has its exception.
    internal ExcelMappingError(ExcelMappingErrorKind kind, ExcelMappingException exception)
    {
        Kind = kind;
        _exception = exception;
        _cell = default;
        _context = default;
        _valueType = null;
        _mapper = null;
        _mapperError = null;
        _emptyHint = null;
    }

    // An empty or invalid cell, whose exception is created by ToException.
    internal ExcelMappingError(ExcelMappingErrorKind kind, ReadCellResult cell, in CellContext context, Type valueType, ICellExceptionSource? mapper, object? mapperError, string emptyHint)
    {
        Kind = kind;
        _exception = null;
        _cell = cell;
        _context = context;
        _valueType = valueType;
        _mapper = mapper;
        _mapperError = mapperError;
        _emptyHint = emptyHint;
    }

    /// <summary>
    /// Gets the kind of problem.
    /// </summary>
    public ExcelMappingErrorKind Kind { get; }

    /// <summary>
    /// Gets the sheet being read.
    /// </summary>
    public ExcelSheet? Sheet => _exception is { } exception ? exception.Sheet : _context.View?.Sheet;

    /// <summary>
    /// Gets the zero-based index of the row in the sheet.
    /// </summary>
    public int RowIndex => _exception?.RowIndex ?? _context.RowIndex;

    /// <summary>
    /// Gets the zero-based index of the cell's column in the sheet, or -1 if the problem isn't with a cell.
    /// </summary>
    public int ColumnIndex => _exception?.ColumnIndex ?? _context.ColumnIndex;

    /// <summary>
    /// Gets the name of the cell's column in the heading, or null if it isn't known.
    /// </summary>
    public string? ColumnName
        => _exception is { } exception ? exception.ColumnName
            : _context.View?.Heading is { } heading && heading.TryGetColumnName(_context.ColumnIndex, out var name) ? name : null;

    /// <summary>
    /// Gets the property or field being read, if known.
    /// </summary>
    public MemberInfo? Member => _exception is { } exception ? exception.Member : _context.Member;

    /// <summary>
    /// Gets the cell that is empty or invalid. For other kinds of problem, it's an empty value.
    /// Reading <see cref="CellValue.Text"/> of a cell that isn't text, such as a number, creates
    /// its text.
    /// </summary>
    public CellValue Value => _exception is null ? new CellValue(_cell) : default;

    /// <summary>
    /// Creates the <see cref="ExcelMappingException"/> that <see cref="ExcelSheet.ReadRows{T}()"/>
    /// would throw for the problem, with its message and the reason the value is invalid.
    /// </summary>
    /// <returns>The exception, which isn't thrown.</returns>
    public ExcelMappingException ToException()
    {
        if (_exception is { } exception)
        {
            return exception;
        }

        var inner = Kind == ExcelMappingErrorKind.InvalidValue ? _mapper?.CreateException(_cell, in _context, _mapperError) : null;
        return CellErrors.Create(new CellValue(_cell), _context, inner, _valueType!, _emptyHint!);
    }
}
