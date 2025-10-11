namespace ExcelMapper;

public interface IColumnIndicesProviderCellReaderFactory
{
    int[]? GetColumnIndices(ExcelSheet sheet);
}
