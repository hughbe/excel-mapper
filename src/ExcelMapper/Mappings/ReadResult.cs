namespace ExcelMapper.Mappings
{
    public struct ReadResult
    {
        public int ColumnIndex { get; }
        public string StringValue { get; }

        public ReadResult(int columnIndex, string stringValue)
        {
            ColumnIndex = columnIndex;
            StringValue = stringValue;
        }
    }
}
