using ExcelMapper.Pipeline.Items;

namespace ExcelMapper.Pipeline
{
    public static class PipelineExtensions
    {
        public static TPipeline DisallowEmptyOrInvalid<TPipeline, T>(this TPipeline pipeline) where TPipeline : Pipeline<T>
        {
            return pipeline
                .WithEmptyFallbackItem(new ThrowIfStatusPipelineItem<T>(PipelineStatus.Empty))
                .WithInvalidFallbackItem(new ThrowIfStatusPipelineItem<T>(PipelineStatus.Invalid));
        }

        public static TPipeline WithEmptyFallback<TPipeline, T>(this TPipeline pipeline, T fallbackValue) where TPipeline : Pipeline<T>
        {
            var fallbackItem = new FixedValuePipelineItem<T>(fallbackValue);
            pipeline.EmptyFallback = fallbackItem;
            return pipeline;
        }

        public static TPipeline WithEmptyFallbackItem<TPipeline, T>(this TPipeline pipeline, PipelineItem<T> fallbackItem) where TPipeline : Pipeline<T>
        {
            pipeline.EmptyFallback = fallbackItem;
            return pipeline;
        }

        public static TPipeline WithInvalidFallback<TPipeline, T>(this TPipeline pipeline, T fallbackValue) where TPipeline : Pipeline<T>
        {
            var fallbackItem = new FixedValuePipelineItem<T>(fallbackValue);
            pipeline.InvalidFallback = fallbackItem;
            return pipeline;
        }

        public static TPipeline WithInvalidFallbackItem<TPipeline, T>(this TPipeline pipeline, PipelineItem<T> fallbackItem) where TPipeline : Pipeline<T>
        {
            pipeline.InvalidFallback = fallbackItem;
            return pipeline;
        }
    }
}
