namespace ExcelMapper.Mappings
{
    public interface ICellValueMapper
    {
        PropertyMapperResultType MapCellValue(ReadCellValueResult mapResult, ref object value);
    }
}
