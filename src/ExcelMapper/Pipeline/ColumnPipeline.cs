using System;
using System.Reflection;

namespace ExcelMapper.Pipeline
{
    public class ColumnPipeline<T> : SinglePipeline<T>
    {
        public string ColumnName { get; }

        public ColumnPipeline(string columnName, MemberInfo member) : base(member)
        {
            if (columnName == null)
            {
                throw new ArgumentNullException(nameof(columnName));
            }

            if (columnName.Length == 0)
            {
                throw new ArgumentException(nameof(columnName));
            }

            ColumnName = columnName;
        }

        protected internal override object Execute(PipelineContext context)
        {
            int index = context.Sheet.Heading.GetColumnIndex(ColumnName);
            context.SetColumnIndex(index);

            return CompletePipeline(context);
        }
    }
}
