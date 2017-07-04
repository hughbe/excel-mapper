using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Pipeline
{
    internal sealed class ColumnsPipeline<T, TElement> : MultiPipeline<T, TElement>
    {
        public string[] ColumnNames { get; }
        private EnumerableType Type { get; }

        internal ColumnsPipeline(string[] columnNames, MemberInfo member) : base(columnNames.Length, member)
        {
            ColumnNames = columnNames;
        }

        protected internal override object Execute(ExcelSheet sheet, ExcelRow row)
        {
            IEnumerable<string> stringValues = ColumnNames.Select(columnName =>
            {
                int index = sheet.Heading.GetColumnIndex(columnName);
                return row.GetString(index);
            });
            return CompletePipeline(stringValues);
        }
    }
}
