using System.Reflection;
using ExcelMapper.Mappings.Support;

namespace ExcelMapper
{
    public class SinglePropertyMapping<T> : SinglePropertyMapping, ISinglePropertyMapping<T>
    {
        public SinglePropertyMapping(MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(member, typeof(T), emptyValueStrategy)
        {
        }
    }
}
