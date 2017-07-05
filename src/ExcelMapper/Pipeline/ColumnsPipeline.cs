using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Pipeline
{
    internal sealed class ColumnsPipeline<T, TElement> : MultiPipeline<T, TElement>
    {
        public string[] ColumnNames { get; }
        private EnumerableType Type { get; }

        internal ColumnsPipeline(string[] columnNames, MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(columnNames.Length, member, emptyValueStrategy)
        {
            ColumnNames = columnNames;
        }

        protected internal override object Execute(PipelineContext context)
        {
            IEnumerable<string> stringValues = ColumnNames.Select(columnName =>
            {
                int index = context.Sheet.Heading.GetColumnIndex(columnName);
                return context.Reader.GetString(index);
            });
            return CompletePipeline(context, stringValues);
        }
    }
}
