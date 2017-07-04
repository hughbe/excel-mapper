using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelMapper
{
    public abstract class ExcelClassMap
    {
        public Type Type { get; }

        private List<Pipeline.Pipeline> Mappings { get; } = new List<Pipeline.Pipeline>();

        internal ExcelClassMap(Type type) => Type = type;

        protected internal void AddMapping(Pipeline.Pipeline pipeline)
        {
            Mappings.Add(pipeline);
        }

        internal object Execute(ExcelSheet sheet, ExcelRow row)
        {
            object value = Activator.CreateInstance(Type);

            foreach (Pipeline.Pipeline pipeline in Mappings)
            {
                pipeline.SetValue(value, sheet, row);
            }

            return value;
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
