using System.Collections.Generic;
using System.Reflection;

namespace ExcelMapper.Pipeline
{
    public abstract class Pipeline<T> : Pipeline
    {
        public Pipeline(MemberInfo member) : base(member) { }

        public List<PipelineItem<T>> Items { get; internal set; } = new List<PipelineItem<T>>();
        public PipelineItem<T> EmptyFallback { get; internal set; }
        public PipelineItem<T> InvalidFallback { get; internal set; }
    }
}
