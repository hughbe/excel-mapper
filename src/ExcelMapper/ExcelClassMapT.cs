using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using ExcelMapper.Mappings.MultiItems;
using ExcelMapper.Utilities;

namespace ExcelMapper
{
    public class ExcelClassMap<T> : ExcelClassMap
    {
        public EmptyValueStrategy EmptyValueStrategy { get; }

        public ExcelClassMap() : base(typeof(T)) { }

        public ExcelClassMap(EmptyValueStrategy emptyValueStrategy) : this()
        {
            if (!Enum.IsDefined(typeof(EmptyValueStrategy), emptyValueStrategy))
            {
                throw new ArgumentException($"Invalid value \"{emptyValueStrategy}\".", nameof(emptyValueStrategy));
            }

            EmptyValueStrategy = emptyValueStrategy;
        }

        public SinglePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);
            MemberInfo member = memberExpression.Member;

            var mapping = new SinglePropertyMapping<TProperty>(member, EmptyValueStrategy);
            AddMapping(mapping);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, IEnumerable<TProperty>>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            return MultiMap<TProperty>(memberExpression);
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, ICollection<TProperty>>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            return MultiMap<TProperty>(memberExpression);
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, IList<TProperty>>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            return MultiMap<TProperty>(memberExpression);
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, List<TProperty>>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            return MultiMap<TProperty>(memberExpression);
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, TProperty[]>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            var mapping = new ArrayMapping<TProperty>(memberExpression.Member, EmptyValueStrategy);
            AddMapping(mapping);
            return mapping;
        }

        private EnumerablePropertyMapping<TProperty> MultiMap<TProperty>(MemberExpression memberExpression)
        {
            var mapping = GetMultiMapping<TProperty>(memberExpression.Member);
            AddMapping(mapping);
            return mapping;
        }

        private EnumerablePropertyMapping<TProperty> GetMultiMapping<TProperty>(MemberInfo member)
        {
            Type type = member.MemberType();
            TypeInfo typeInfo = type.GetTypeInfo();

            if (typeInfo.IsInterface)
            {
                if (type.IsAssignableFrom(typeof(List<TProperty>)))
                {
                    return new InterfaceAssignableFromListMapping<TProperty>(member, EmptyValueStrategy);
                }
            }
            else if (type.ImplementsInterface(typeof(ICollection<TProperty>)))
            {
                return new ConcreteICollectionMapping<TProperty>(type, member, EmptyValueStrategy);
            }

            throw new ExcelMappingException($"No known way to instantiate type \"{type}\". It must be an array, be assignable from List<T> or implement ICollection<T>.");
        }
    }
}
