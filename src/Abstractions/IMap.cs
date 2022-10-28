
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Abstractions;

public interface IMap
{
    bool TryMap(ExcelRow row, IExcelDataReader reader, MemberInfo member, out object value);
}
