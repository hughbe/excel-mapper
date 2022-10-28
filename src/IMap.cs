using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper;
 
public interface IMap
{
    bool TryGetValue(ExcelRow row, IExcelDataReader reader, MemberInfo member, out object value);
}
