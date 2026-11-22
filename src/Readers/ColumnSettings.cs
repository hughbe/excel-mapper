using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Where a map reads its cells, and how: its <see cref="Readers.Columns"/>, whether they can be
/// missing, and whether cells are read as the text Excel shows. Every map that reads its own
/// columns (<see cref="IColumnsMap"/>) has one, and it can't change once the map is frozen.
/// </summary>
internal sealed class ColumnSettings(Columns columns)
{
    private Columns _columns = columns;
    private bool _optional;
    private bool _preserveFormatting;
    private bool _frozen;

    /// <summary>
    /// Gets or sets the columns the map reads.
    /// </summary>
    public Columns Columns
    {
        get => _columns;
        set
        {
            ThrowIfFrozen();
            _columns = value;
        }
    }

    /// <summary>
    /// Gets or sets whether a missing column leaves the member's value, rather than being an error.
    /// </summary>
    public bool Optional
    {
        get => _optional;
        set
        {
            ThrowIfFrozen();
            _optional = value;
        }
    }

    /// <summary>
    /// Gets or sets whether cells are read as the text Excel shows, with their number format applied.
    /// </summary>
    public bool PreserveFormatting
    {
        get => _preserveFormatting;
        set
        {
            ThrowIfFrozen();
            _preserveFormatting = value;
        }
    }

    /// <summary>
    /// Stops the settings changing, when their map is frozen (see <see cref="IMap.Freeze"/>).
    /// </summary>
    public void Freeze() => _frozen = true;

    /// <summary>
    /// Throws if the map has been used to read rows.
    /// </summary>
    public void ThrowIfFrozen() => ThrowHelpers.ThrowIfFrozen(_frozen);

    /// <summary>
    /// What reading a member whose column is missing gives: no value (false) if the map is
    /// optional, otherwise an <see cref="ExcelMappingException"/> naming the column.
    /// </summary>
    public bool MissingColumn(SheetView sheet, int rowIndex, MemberInfo? member)
        => Optional ? false : throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _columns, member);

    /// <summary>
    /// Starts reading the cells of a collection or dictionary from the current row.
    /// </summary>
    /// <returns>False if the cells are missing and the map is optional.</returns>
    /// <exception cref="ExcelMappingException">The cells are missing and the map isn't optional.</exception>
    public bool TryStart([NotNullWhen(true)] ICellsReader? cellsReader, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, out int count)
    {
        if (cellsReader is not null && cellsReader.Start(reader, _preserveFormatting, out count))
        {
            return true;
        }

        count = 0;
        return MissingColumn(sheet, rowIndex, member);
    }
}
