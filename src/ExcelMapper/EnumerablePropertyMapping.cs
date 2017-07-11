using System.Reflection;

namespace ExcelMapper
{
    public abstract class EnumerablePropertyMapping : PropertyMapping
    {
        public EnumerablePropertyMapping(MemberInfo member) : base(member)
        {
        }
    }
}
