using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelMapper.Mappings.MultiItems
{
    internal class ConcreteICollectionMapping<T> : EnumerablePropertyMapping<T>
    {
        public Type CollectionType { get; }

        public ConcreteICollectionMapping(Type type, MemberInfo member, SinglePropertyMapping<T> elementMapping) : base(member, elementMapping)
        {
            CollectionType = type;
        }

        public override object CreateFromElements(IEnumerable<T> elements)
        {
            ICollection<T> value = (ICollection<T>)Activator.CreateInstance(CollectionType);

            foreach (T element in elements)
            {
                value.Add(element);
            }

            return value;
        }
    }
}
