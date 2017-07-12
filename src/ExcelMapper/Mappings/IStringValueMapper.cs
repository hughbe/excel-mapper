namespace ExcelMapper.Mappings
{
    public interface IStringValueMapper
    {
        PropertyMappingResultType GetProperty(ReadResult mapResult, ref object value);
    }
}
