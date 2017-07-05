using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Pipeline
{
    internal sealed class IndicesPipeline<T, TElement> : MultiPipeline<T, TElement>
    {
        public int[] Indices { get; }
        private EnumerableType Type { get; }

        internal IndicesPipeline(int[] indices, MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(indices.Length, member, emptyValueStrategy)
        {
            Indices = indices;
        }

        protected internal override object Execute(PipelineContext context)
        {
            IEnumerable<string> stringValues = Indices.Select(index => context.Reader.GetString(index));
            return CompletePipeline(context, stringValues);
        }
    }
}
