using ExcelDataReader;

namespace ExcelMapper
{
    public class ExcelRow
    {
        internal ExcelRow(int index, IExcelDataReader reader)
        {
            Index = index;
            Reader = reader;
        }

        public int Index { get; }
        private IExcelDataReader Reader { get; }

        internal string GetString(int index) => Reader.GetString(index);
    }
}
