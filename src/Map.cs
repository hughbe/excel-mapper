using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    public abstract class Map
    {
        public abstract bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo member, out object value);
    }
}
