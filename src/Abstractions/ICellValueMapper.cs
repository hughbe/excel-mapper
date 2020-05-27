namespace ExcelMapper.Abstractions
{
    public interface ICellValueMapper
    {
        PropertyMapperResultType MapCellValue(ReadCellValueResult mapResult, ref object value);
    }
}
