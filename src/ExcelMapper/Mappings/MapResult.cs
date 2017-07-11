namespace ExcelMapper.Mappings
{
    public struct MapResult
    {
        public int ColumnIndex { get; }
        public string StringValue { get; set; }

        public MapResult(int index, string stringValue)
        {
            ColumnIndex = index;
            StringValue = stringValue;
        }
    }
}
