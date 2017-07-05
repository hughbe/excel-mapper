using System;
using System.Reflection;

namespace ExcelMapper.Pipeline
{
    public abstract class Pipeline
    {
        public MemberInfo Member { get; }

        public Pipeline(MemberInfo member)
        {
            Member = member ?? throw new ArgumentNullException(nameof(member));
        }

        internal void SetValue(object value, PipelineContext context)
        {
            object propertyValue = Execute(context);
            if (Member is FieldInfo field)
            {
                field.SetValue(value, propertyValue);
            }
            else if (Member is PropertyInfo property)
            {
                property.SetValue(value, propertyValue);
            }
            else
            {
                throw new ExcelMappingException("Unknown member.");
            }
        }

        protected internal abstract object Execute(PipelineContext context);
    }
}
