namespace ExcelMapper.Mappings
{
    public interface ISinglePropertyMappingItem
    {
        PropertyMappingResult GetProperty(ReadResult mapResult);
    }
}
