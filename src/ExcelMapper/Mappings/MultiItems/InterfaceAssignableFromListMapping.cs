using System.Collections.Generic;
using System.Reflection;

namespace ExcelMapper.Mappings.MultiItems
{
    internal class InterfaceAssignableFromListMapping<T> : EnumerablePropertyMapping<T>
    {
        public InterfaceAssignableFromListMapping(MemberInfo member, SinglePropertyMapping<T> elementMapping) : base(member, elementMapping)
        {
        }

        public override object CreateFromElements(IEnumerable<T> elements) => elements;
    }
}
