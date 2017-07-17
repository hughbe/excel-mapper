namespace ExcelMapper.Mappings
{
    public interface ICellValueMapper
    {
        PropertyMappingResultType GetProperty(ReadCellValueResult mapResult, ref object value);
    }
}
