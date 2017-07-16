using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Mappings.MultiItems
{
    internal class ArrayMapping<T> : EnumerablePropertyMapping<T>
    {
        public ArrayMapping(MemberInfo member, SinglePropertyMapping<T> elementMapping) : base(member, elementMapping)
        {
        }

        public override object CreateFromElements(IEnumerable<T> elements) => elements.ToArray();
    }
}
