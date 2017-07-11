using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using ExcelMapper.Mappings.MultiItems;

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

            var mapping = new AssignableFromListMapping<TProperty>(memberExpression.Member, EmptyValueStrategy);
            AddMapping(mapping);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, ICollection<TProperty>>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            var mapping = new AssignableFromListMapping<TProperty>(memberExpression.Member, EmptyValueStrategy);
            AddMapping(mapping);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, IList<TProperty>>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            var mapping = new AssignableFromListMapping<TProperty>(memberExpression.Member, EmptyValueStrategy);
            AddMapping(mapping);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, TProperty[]>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            var mapping = new ArrayMapping<TProperty>(memberExpression.Member, EmptyValueStrategy);
            AddMapping(mapping);
            return mapping;
        }
    }
}
