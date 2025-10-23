namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell and produces multiple values by splitting the string value
/// using the given separators.
/// </summary>
public class CharSplitReaderFactory : SplitReaderFactory
{
    private char[] _separators = [','];

    /// <summary>
    /// Gets or sets the separators used to split the string value of the cell.
    /// </summary>
    public char[] Separators
    {
        get => _separators;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            if (value.Length == 0)
            {
                throw new ArgumentException("Separators cannot be empty.", nameof(value));
            }

            _separators = value;
        }
    }

    /// <summary>
    /// Constructs a reader that reads the string value of a cell and produces multiple values
    /// by splitting it.
    /// </summary>
    /// <param name="cellReader">The ICellReaderFactory that reads the string value of the cell before it is split.</param>
    public CharSplitReaderFactory(ICellReaderFactory cellReader) : base(cellReader)
    {
    }

    /// <inheritdoc/>
    protected override string[] GetValues(string value) => value.Split(Separators, Options);
}
