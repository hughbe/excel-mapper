namespace ExcelMapper;

public interface IColumnNameProviderCellReaderFactory
{
    string GetColumnName(ExcelSheet sheet);
}
