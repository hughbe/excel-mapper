using System.Reflection;

namespace ExcelMapper.Pipeline
{
    public abstract class SinglePipeline<T> : Pipeline<T>
    {
        public SinglePipeline(MemberInfo member) : base(member)
        {
        }

        protected object CompletePipeline(string stringValue)
        {
            PipelineResult<T> result = new PipelineResult<T>(PipelineStatus.Began, stringValue, default(T));
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
