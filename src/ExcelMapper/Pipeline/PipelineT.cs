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

        protected T CompletePipeline(PipelineContext context)
        {
            PipelineResult<T> result = new PipelineResult<T>(PipelineStatus.Began, context, default(T));
            for (int i = 0; i < Items.Count; i++)
            {
                PipelineItem<T> item = Items[i];
                result = item.TryMap(result);
                if (result.Status == PipelineStatus.Completed)
                {
                    return result.Result;
                }
            }

            if (result.Status == PipelineStatus.Empty && EmptyFallback != null)
            {
                result = EmptyFallback.TryMap(result);
            }

            if (result.Status == PipelineStatus.Invalid && InvalidFallback != null)
            {
                result = InvalidFallback.TryMap(result);
            }

            return result.Result;
        }
    }
}
