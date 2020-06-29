using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    public interface IMap
    {
        bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo member, out object value);
    }
}
