using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Mappings.MultiItems
{
    /// <summary>
    /// Describes a property map that maps multiple values of one or multiple cells to T[].
    /// </summary>
    /// <typeparam name="T">The element type of the array to create.</typeparam>
    internal class ArrayPropertyMap<T> : EnumerableExcelPropertyMap<T>
    {
        public ArrayPropertyMap(MemberInfo member, OneToOnePropertyMap<T> elementMapping) : base(member, elementMapping)
        {
        }

        protected override object CreateFromElements(IEnumerable<T> elements) => elements.ToArray();
    }
}
