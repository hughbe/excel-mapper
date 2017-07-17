namespace ExcelMapper.Mappings.Mappers
{
    /// <summary>
    /// A mapper that returns the string value of a cell.
    /// </summary>
    public class StringMapper : ICellValueMapper
    {
        public PropertyMappingResultType GetProperty(ReadCellValueResult result, ref object value)
        {
            value = result.StringValue;
            return PropertyMappingResultType.SuccessIfNoOtherSuccess;
        }
    }
}
