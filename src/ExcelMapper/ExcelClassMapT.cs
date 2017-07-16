using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
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
            MemberExpression memberExpression = GetMemberExpression(expression);
            MemberInfo member = memberExpression.Member;

            bool canMap = member.AutoMap(EmptyValueStrategy, out SinglePropertyMapping<TProperty> mapping);
            if (!canMap)
            {
                throw new ExcelMappingException($"Don't know how to map type {typeof(TProperty)}.");
            }

            AddMapping(mapping, expression);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, IEnumerable<TProperty>>> expression)
        {
            MemberExpression memberExpression = GetMemberExpression(expression);
            var mapping = GetMultiMapping<TProperty>(memberExpression.Member);

            AddMapping(mapping, expression);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, ICollection<TProperty>>> expression)
        {
            MemberExpression memberExpression = GetMemberExpression(expression);
            var mapping = GetMultiMapping<TProperty>(memberExpression.Member);

            AddMapping(mapping, expression);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, IList<TProperty>>> expression)
        {
            MemberExpression memberExpression = GetMemberExpression(expression);
            var mapping = GetMultiMapping<TProperty>(memberExpression.Member);

            AddMapping(mapping, expression);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, List<TProperty>>> expression)
        {
            MemberExpression memberExpression = GetMemberExpression(expression);
            var mapping = GetMultiMapping<TProperty>(memberExpression.Member);

            AddMapping(mapping, expression);
            return mapping;
        }

        public EnumerablePropertyMapping<TProperty> Map<TProperty>(Expression<Func<T, TProperty[]>> expression)
        {
            MemberExpression memberExpression = GetMemberExpression(expression);
            var mapping = GetMultiMapping<TProperty>(memberExpression.Member);

            AddMapping(mapping, expression);
            return mapping;
        }

        public ObjectPropertyMapping<TProperty> MapObject<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            MemberExpression memberExpression = GetMemberExpression(expression);
            MemberInfo member = memberExpression.Member;

            if (!member.AutoMapObject(EmptyValueStrategy, out ObjectPropertyMapping<TProperty> mapping))
            {
                throw new ExcelMappingException($"Could not map object of type \"{typeof(TProperty)}\".");
            }

            AddMapping(mapping, expression);
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

        protected internal MemberExpression GetMemberExpression<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            if (!(expression.Body is MemberExpression rootMemberExpression))
            {
                throw new ArgumentException("Not a member expression.", nameof(expression));
            }

            return rootMemberExpression;
        }

        protected internal void AddMapping<TProperty>(PropertyMapping mapping, Expression<Func<T, TProperty>> expression)
        {
            Expression expressionBody = expression.Body;
            var expressions = new Stack<MemberExpression>();
            while (expressionBody != null)
            {
                if (!(expressionBody is MemberExpression memberExpressionBody))
                {
                    // Each mapping is of the form (parameter => member).
                    if (expressionBody is ParameterExpression parameterExpression)
                    {
                        break;
                    }

                    throw new ArgumentException($"Expression can only contain member accesses, but found {expressionBody}.", nameof(expression));
                }
                
                expressions.Push(memberExpressionBody);
                expressionBody = memberExpressionBody.Expression;
            }

            if (expressions.Count == 1)
            {
                // Simple case: parameter => prop
                Mappings.Add(mapping);
            }
            else
            {
                // Go through the chain of members and make sure that they are valid.
                // E.g. parameter => parameter.prop.subprop.field.
                CreateObjectMap(mapping, expressions);
            }
        }
    }
}
