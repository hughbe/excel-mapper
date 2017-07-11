using System.Collections.Generic;
using System.Reflection;

namespace ExcelMapper.Mappings.MultiItems
{
    internal class AssignableFromListMapping<T> : EnumerablePropertyMapping<T>
    {
        public AssignableFromListMapping(MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(member, emptyValueStrategy)
        {
        }

        public override object CreateFromElements(IEnumerable<T> elements) => elements;
    }
}
