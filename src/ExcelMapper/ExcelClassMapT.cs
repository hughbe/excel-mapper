using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelMapper.Pipeline;
using ExcelMapper.Utilities;

namespace ExcelMapper
{
    public class ExcelClassMap<T> : ExcelClassMap where T : new()
    {
        public ExcelClassMap() : base(typeof(T)) { }

        public DefaultPipeline<TProperty> Map<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            MemberExpression memberExpression = ValidateExpression(expression);

            var propertyMap = new DefaultPipeline<TProperty>(memberExpression.Member);
            AddMapping(propertyMap);
            return propertyMap;
        }

        public MultiPipeline<TProperty, TElement> MultiMap<TProperty, TElement>(Expression<Func<T, TProperty>> expression, params string[] columnNames)
        {
            if (columnNames == null)
            {
                throw new ArgumentNullException(nameof(columnNames));
            }

            if (columnNames.Length == 0)
            {
                throw new ArgumentException("Column names cannot be empty", nameof(columnNames));
            }

            MemberExpression memberExpression = ValidateExpression(expression);
            Type elementType = memberExpression.Type.GetIEnumerableType();
            if (elementType == null)
            {
                throw new ExcelMappingException($"Cannot map type {typeof(TProperty)}. It does not implement IEnumerable<T>.");
            }

            Type mapType = typeof(ColumnsPipeline<,>).MakeGenericType(typeof(TProperty), elementType);
            ConstructorInfo constructor = mapType.GetConstructor(new Type[] { typeof(string[]), typeof(MemberInfo) });

            MultiPipeline<TProperty, TElement> propertyMap = (MultiPipeline<TProperty, TElement>)constructor.Invoke(new object[] { columnNames, memberExpression.Member });
            AddMapping(propertyMap);
            return propertyMap;
        }

        public MultiPipeline<TProperty, TElement> MultiMap<TProperty, TElement>(Expression<Func<T, TProperty>> expression, IEnumerable<string> columnNames)
        {
            return MultiMap<TProperty, TElement>(expression, columnNames?.ToArray());
        }

        public MultiPipeline<TProperty, TElement> MultiMap<TProperty, TElement>(Expression<Func<T, TProperty>> expression, params int[] indices)
        {
            if (indices == null)
            {
                throw new ArgumentNullException(nameof(indices));
            }

            if (indices.Length == 0)
            {
                throw new ArgumentException("Indices cannot be empty", nameof(indices));
            }

            MemberExpression memberExpression = ValidateExpression(expression);
            Type elementType = memberExpression.Type.GetIEnumerableType();
            if (elementType == null)
            {
                throw new ExcelMappingException($"Cannot map type {typeof(TProperty)}. It does not implement IEnumerable<T>.");
            }

            Type mapType = typeof(IndicesPipeline<,>).MakeGenericType(typeof(TProperty), elementType);
            ConstructorInfo constructor = mapType.GetConstructor(new Type[] { typeof(int[]), typeof(MemberInfo) });

            MultiPipeline<TProperty, TElement> propertyMap = (MultiPipeline<TProperty, TElement>)constructor.Invoke(new object[] { indices, memberExpression.Member });
            AddMapping(propertyMap);
            return propertyMap;
        }

        public MultiPipeline<TProperty, TElement> MultiMap<TProperty, TElement>(Expression<Func< TProperty>> expression, IEnumerable<int> indices)
        {
            return MultiMap<TProperty, TElement>(expression, indices?.ToArray());
        }
    }
}
