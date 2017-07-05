using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelMapper.Pipeline;

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

        internal object Execute(PipelineContext context)
        {
            object value = Activator.CreateInstance(Type);

            foreach (Pipeline.Pipeline pipeline in Mappings)
            {
                pipeline.SetValue(value, context);
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
