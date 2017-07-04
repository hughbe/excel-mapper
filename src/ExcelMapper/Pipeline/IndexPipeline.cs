using System;
using System.Reflection;

namespace ExcelMapper.Pipeline
{
    public sealed class IndexPipeline<T> : SinglePipeline<T>
    {
        public int Index { get; }

        public IndexPipeline(int index, MemberInfo member) : base(member)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            Index = index;
        }

        protected internal override object Execute(ExcelSheet sheet, ExcelRow row)
        {
            string stringValue = row.GetString(Index);
            return CompletePipeline(stringValue);
        }
    }
}
