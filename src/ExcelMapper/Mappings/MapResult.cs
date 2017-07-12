namespace ExcelMapper.Mappings
{
    public struct MapResult
    {
        public int ColumnIndex { get; }
        public string StringValue { get; }

        public MapResult(int coolumnIndex, string stringValue)
        {
            ColumnIndex = coolumnIndex;
            StringValue = stringValue;
        }
    }
}
