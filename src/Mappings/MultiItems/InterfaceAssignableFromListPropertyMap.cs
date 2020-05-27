using System.Collections.Generic;
using System.Reflection;

namespace ExcelMapper.Mappings.MultiItems
{
    /// <summary>
    /// Describes a property map that maps multiple values of one or multiple cells to an interface that
    /// is assignable from List&lt;&gt;.
    /// </summary>
    /// <typeparam name="T">The element type of the List to create.</typeparam>
    internal class InterfaceAssignableFromListPropertyMap<T> : EnumerableExcelPropertyMap<T>
    {
        public InterfaceAssignableFromListPropertyMap(MemberInfo member, ValuePipeline elementMapping) : base(member, elementMapping)
        {
        }

        protected override object CreateFromElements(IEnumerable<T> elements) => elements;
    }
}
