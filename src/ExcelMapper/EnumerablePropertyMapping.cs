using System.Reflection;

namespace ExcelMapper
{
    public abstract class EnumerablePropertyMapping : PropertyMapping
    {
        internal EnumerablePropertyMapping(MemberInfo member) : base(member)
        {
        }
    }
}
