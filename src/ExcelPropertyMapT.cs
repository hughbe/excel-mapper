using System.Reflection;

namespace ExcelMapper
{
    public class ExcelPropertyMap<T> : ExcelPropertyMap
    {
        public ExcelPropertyMap(MemberInfo member, IMap map) : base(member, map)
        {
        }
    }
}
