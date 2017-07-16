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

            bool canMap = member.AutoMap(EmptyValueStrategy, out SinglePropertyMapping<TProperty> mapping);
            if (!canMap)
            {
                throw new ExcelMappingException($"Don't know how to map type {typeof(TProperty)}.");
            }

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
            return MultiMap<TProperty>(memberExpression);
        }

        public ObjectPropertyMapping<TProperty> MapObject<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);
            MemberInfo member = memberExpression.Member;

            if (!member.AutoMapObject(EmptyValueStrategy, out ObjectPropertyMapping<TProperty> mapping))
            {
                throw new ExcelMappingException($"Could not map object of type \"{typeof(TProperty)}\".");
            }

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
            if (!member.AutoMapEnumerable(EmptyValueStrategy, out EnumerablePropertyMapping<TProperty> mapping))
            {
                throw new ExcelMappingException($"No known way to instantiate type \"{typeof(TProperty)}\". It must be a single dimensional array, be assignable from List<T> or implement ICollection<T>.");
            }

            return mapping;
        }
    }
}
