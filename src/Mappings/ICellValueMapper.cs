namespace ExcelMapper.Mappings
{
    public interface ICellValueMapper
    {
        PropertyMapperResultType GetProperty(ReadCellValueResult mapResult, ref object value);
    }
}
