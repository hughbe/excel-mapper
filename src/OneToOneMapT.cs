using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper;
 
public class OneToOneMap<T> : IOneToOneMap<T>, IMap<T>, IMap
{
    public OneToOneMap(ICellReader reader)
    {
        Reader = reader ?? throw new ArgumentNullException(nameof(reader));
    }

    private ICellReader _reader;

    /// <summary>
    /// Gets or sets the <see cref="ICellReader"/> used to read the cell value.
    /// </summary>
    public ICellReader Reader
    {
        get => _reader;
        set => _reader = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// Gets or sets a value indicating whether the cell value is optional.
    /// </summary>
    public bool Optional { get; set; }

    /// <summary>
    /// Gets the list of <see cref="ICellMapper"/>s used to convert the cell value.
    /// </summary>
    public CellValueMapperCollection Mappers { get; } = new CellValueMapperCollection();

    public bool TryMap(ExcelRow row, IExcelDataReader reader, MemberInfo member, out object result)
    {
        if (!Reader.TryGetCell(row, out ExcelCell cell))
        {
            if (Optional)
            {
                result = default;
                return false;
            }

            throw new ExcelMappingException($"Could not read value for {member.Name}", row.Sheet, row.RowIndex, -1);
        }

        var cellValue = reader.GetValue(cell.ColumnIndex);
        result = (T)ValuePipeline.GetPropertyValue(cell, cellValue, member, Mappers);
        return true;
    }
}
