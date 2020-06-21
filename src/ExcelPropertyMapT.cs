using System.Reflection;

namespace ExcelMapper
{
    public class ExcelPropertyMap<T> : ExcelPropertyMap
    {
        public ExcelPropertyMap(MemberInfo member, Map map) : base(member, map)
        {
        }
    }
}
