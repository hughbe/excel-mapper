namespace ExcelMapper.Abstractions
{
    public interface ICellValueMapper
    {
        CellValueMapperResult MapCellValue(ReadCellValueResult readResult);
    }
}
