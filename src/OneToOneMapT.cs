using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper;
 
public class OneToOneMap<T> : IMapper<T>, IMap
{
    public OneToOneMap(ICellReader reader)
    {
        CellReader = reader ?? throw new ArgumentNullException(nameof(reader));
    }

    private ICellReader _reader;

    public ICellReader CellReader
    {
        get => _reader;
        set => _reader = value ?? throw new ArgumentNullException(nameof(value));
    }

    public bool Optional { get; set; }

    public CellValueMapperCollection Mappers { get; } = new CellValueMapperCollection();

    public bool TryGetValue(ExcelRow row, IExcelDataReader reader, MemberInfo member, out object result)
    {
        if (!CellReader.TryGetCell(row, out ExcelCell cell))
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
