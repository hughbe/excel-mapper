using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelDataReader;

namespace ExcelMapper
{
    public abstract class ExcelClassMap
    {
        public Type Type { get; }

        private List<PropertyMapping> Mappings { get; } = new List<PropertyMapping>();

        internal ExcelClassMap(Type type) => Type = type;

        protected internal void AddMapping(PropertyMapping mapping)
        {
            Mappings.Add(mapping);
        }

        internal object Execute(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            object instance = Activator.CreateInstance(Type);

            foreach (PropertyMapping pipeline in Mappings)
            {
                object propertyValue = pipeline.GetPropertyValue(sheet, rowIndex, reader);
                pipeline.SetPropertyFactory(instance, propertyValue);
            }

            return instance;
        }

        protected internal static MemberExpression ValidateExpression<T, TProperty>(Expression<Func<T, TProperty>> expression)
        {
            if (!(expression.Body is MemberExpression memberExpression))
            {
                throw new InvalidOperationException("Not a member expression.");
            }

            return memberExpression;
        }
    }
}
