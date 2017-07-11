using System.Collections.Generic;
using System.Reflection;

namespace ExcelMapper.Mappings.MultiItems
{
    internal class InterfaceAssignableFromListMapping<T> : EnumerablePropertyMapping<T>
    {
        public InterfaceAssignableFromListMapping(MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(member, emptyValueStrategy)
        {
        }

        public override object CreateFromElements(IEnumerable<T> elements) => elements;
    }
}
