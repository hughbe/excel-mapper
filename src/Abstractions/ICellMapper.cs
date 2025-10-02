namespace ExcelMapper.Abstractions;

public interface ICellMapper
{
    CellMapperResult MapCellValue(ReadCellResult readResult);
}
