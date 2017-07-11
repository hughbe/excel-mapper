using System.Collections.Generic;
using System.Reflection;

namespace ExcelMapper.Mappings.MultiItems
{
    internal class IEnumerableMapping<T> : MultiPropertyMapping<T>
    {
        public IEnumerableMapping(MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(member, emptyValueStrategy)
        {
        }

        public override object CreateFromElements(IEnumerable<T> elements) => elements;
    }
}
