namespace ExcelMapper.Abstractions;

public interface IExcelColumnMatcher
{
    bool ColumnMatches(ExcelSheet sheet, int columnIndex);
}
