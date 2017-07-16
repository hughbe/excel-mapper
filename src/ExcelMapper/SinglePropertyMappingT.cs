using System.Reflection;
using ExcelMapper.Mappings.Support;

namespace ExcelMapper
{
    public class SinglePropertyMapping<T> : SinglePropertyMapping, ISinglePropertyMapping<T>
    {
        public SinglePropertyMapping(MemberInfo member) : base(member, typeof(T))
        {
        }
    }
}
