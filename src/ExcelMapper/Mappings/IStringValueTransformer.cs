using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public interface IStringValueTransformer
    {
        string TransformStringValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue);
    }
}
