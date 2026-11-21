using ExcelMapper.Abstractions;

namespace ExcelMapper;

/// <summary>
/// The value of a cell, as passed to an <see cref="ICellConverter{T}"/>. Numbers, booleans, dates
/// and durations are read without converting them to text.
/// </summary>
public readonly struct CellValue
{
    private readonly ReadCellResult _cell;

    internal CellValue(ReadCellResult cell)
    {
        // Read the value now, so that it doesn't change when the reader moves to the next row.
        cell.GetValue();
        if (cell.PreserveFormatting)
        {
            cell.GetString();
        }

        _cell = cell;
    }

    /// <summary>
    /// Creates a cell value holding text, for example to test a converter.
    /// </summary>
    /// <param name="text">The text. Null or empty text is an empty cell.</param>
    public CellValue(string? text)
    {
        _cell = new ReadCellResult(0, text, preserveFormatting: false);
    }

    /// <summary>
    /// Creates a cell value holding a number, as Excel stores numbers.
    /// </summary>
    public CellValue(double number)
    {
        _cell = new ReadCellResult(0, (object)number);
    }

    /// <summary>
    /// Creates a cell value holding <see langword="true"/> or <see langword="false"/>.
    /// </summary>
    public CellValue(bool boolean)
    {
        _cell = new ReadCellResult(0, (object)boolean);
    }

    /// <summary>
    /// Creates a cell value holding a date, or a date and time.
    /// </summary>
    public CellValue(DateTime dateTime)
    {
        _cell = new ReadCellResult(0, (object)dateTime);
    }

    /// <summary>
    /// Creates a cell value holding a duration.
    /// </summary>
    public CellValue(TimeSpan timeSpan)
    {
        _cell = new ReadCellResult(0, (object)timeSpan);
    }

    /// <summary>
    /// The cell this value came from, or a cell holding the value for <paramref name="columnIndex"/>
    /// if it was created, for example by a transformer.
    /// </summary>
    internal ReadCellResult ToReadCellResult(int columnIndex)
        => _cell.Reader is not null || _cell.ColumnIndex == columnIndex ? _cell : new ReadCellResult(columnIndex, _cell.GetValue());

    /// <summary>
    /// Gets the kind of value the cell holds. With formatting preserved (see
    /// <see cref="ExcelPreserveFormattingAttribute"/>), every cell that isn't empty is
    /// <see cref="CellKind.Text"/>, as it's read as the text Excel shows.
    /// </summary>
    public CellKind Kind
    {
        get
        {
            if (IsEmpty)
            {
                return CellKind.Empty;
            }

            if (_cell.PreserveFormatting)
            {
                return CellKind.Text;
            }

            return _cell.GetValue() switch
            {
                string => CellKind.Text,
                double => CellKind.Number,
                bool => CellKind.Boolean,
                System.DateTime => CellKind.DateTime,
                System.TimeSpan => CellKind.TimeSpan,
                _ => CellKind.Other,
            };
        }
    }

    /// <summary>
    /// Gets whether the cell is empty or holds empty text.
    /// </summary>
    public bool IsEmpty => _cell.IsEmpty();

    /// <summary>
    /// Gets the cell's value as text: the text of a text cell, the text Excel shows if formatting is
    /// preserved, and otherwise the value converted to a string. Null for an empty cell.
    /// </summary>
    public string? Text => _cell.GetString();

    /// <summary>
    /// Gets the cell's value as Excel stores it: a <see cref="string"/>, <see cref="double"/>,
    /// <see cref="bool"/>, <see cref="System.DateTime"/> or <see cref="System.TimeSpan"/>, or null
    /// for an empty cell. Reading it doesn't box the value again.
    /// </summary>
    public object? Value => _cell.PreserveFormatting ? _cell.GetString() : _cell.GetValue();

    /// <summary>
    /// Gets the number in a <see cref="CellKind.Number"/> cell.
    /// </summary>
    /// <exception cref="InvalidOperationException">The cell doesn't hold a number.</exception>
    public double Number => Kind == CellKind.Number ? (double)_cell.GetValue()! : throw NotKind(CellKind.Number);

    /// <summary>
    /// Gets the value of a <see cref="CellKind.Boolean"/> cell.
    /// </summary>
    /// <exception cref="InvalidOperationException">The cell doesn't hold a boolean.</exception>
    public bool Boolean => Kind == CellKind.Boolean ? (bool)_cell.GetValue()! : throw NotKind(CellKind.Boolean);

    /// <summary>
    /// Gets the date of a <see cref="CellKind.DateTime"/> cell.
    /// </summary>
    /// <exception cref="InvalidOperationException">The cell doesn't hold a date.</exception>
    public DateTime DateTime => Kind == CellKind.DateTime ? (DateTime)_cell.GetValue()! : throw NotKind(CellKind.DateTime);

    /// <summary>
    /// Gets the duration of a <see cref="CellKind.TimeSpan"/> cell.
    /// </summary>
    /// <exception cref="InvalidOperationException">The cell doesn't hold a duration.</exception>
    public TimeSpan TimeSpan => Kind == CellKind.TimeSpan ? (TimeSpan)_cell.GetValue()! : throw NotKind(CellKind.TimeSpan);

    /// <summary>
    /// Returns <see cref="Text"/>, or an empty string for an empty cell.
    /// </summary>
    public override string ToString() => Text ?? string.Empty;

    private InvalidOperationException NotKind(CellKind expected)
        => new($"The cell holds {Kind}, not {expected}.");
}
