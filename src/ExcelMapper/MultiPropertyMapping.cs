using System.Reflection;
using ExcelMapper.Mappings;

namespace ExcelMapper
{
    public abstract class MultiPropertyMapping : PropertyMapping
    {
        public IMultiPropertyMapper Mapper { get; internal set; }

        internal MultiPropertyMapping(MemberInfo member) : base(member)
        {
        }
    }
}
