using ExcelDataReader;

namespace ExcelMapper.Pipeline
{
    public class PipelineContext
    {
        public ExcelSheet Sheet { get; }
        public int  RowIndex { get; }
        public IExcelDataReader Reader { get; }

        public int ColumnIndex { get; private set; }
        public string StringValue { get; set; }

        internal PipelineContext(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            Sheet = sheet;
            RowIndex = rowIndex;
            Reader = reader;
        }

        internal void SetColumnIndex(int index)
        {
            ColumnIndex = index;
            StringValue = Reader.GetString(index);
        }
    }
}
