namespace ExcelMapper;

public interface IColumnNamesProviderCellReaderFactory
{
    string[]? GetColumnNames(ExcelSheet sheet);
}
