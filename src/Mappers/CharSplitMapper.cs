#if MULTI
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;
 
/// <summary>
/// Reads the value of a cell and produces multiple values by splitting the string value
/// using the given separators.
/// </summary>
public class CharSplitCellValueReader : SplitCellValueReader
{
    private char[] _separators = new char[] { ',' };

    /// <summary>
    /// Gets or sets the separators used to split the string value of the cell.
    /// </summary>
    public char[] Separators
    {
        get => _separators;
        set
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            if (value.Length == 0)
            {
                throw new ArgumentException("Separators cannot be empty.", nameof(value));
            }

            _separators = value;
        }
    }

    protected override string[] GetValues(string value) => value.Split(Separators, Options);
}
#endif
