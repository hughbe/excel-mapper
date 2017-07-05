using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Pipeline.Items;

namespace ExcelMapper.Pipeline
{
    public static class SinglePipelineExtensions
    {
        public static TPipeline WithMapping<TPipeline, T>(this TPipeline pipeline, IDictionary<string, T> mapping, IEqualityComparer<string> comparer = null) where TPipeline : SinglePipeline<T>
        {
            var item = new MapStringValuePipelineItem<T>(mapping, comparer);
            pipeline.Items.Add(item);
            return pipeline;
        }

        public static TPipeline WithConverter<TPipeline, T>(this TPipeline pipeline, ConvertUsingSimple<T> mapping) where TPipeline : SinglePipeline<T>
        {
            var item = new ConvertUsingPipelineItem<T>(mapping);
            pipeline.Items.Add(item);
            return pipeline;
        }

        public static TPipeline WithAdditionalDateFormats<TPipeline>(this TPipeline pipeline, params string[] formats) where TPipeline : SinglePipeline<DateTime>
        {
            return pipeline.WithAdditionalDateFormats((IEnumerable<string>)formats);
        }

        public static TPipeline WithAdditionalDateFormats<TPipeline>(this TPipeline pipeline, IEnumerable<string> formats) where TPipeline : SinglePipeline<DateTime>
        {
            ParseAsDateTimePipelineItem dateTimeItem = (ParseAsDateTimePipelineItem)pipeline.Items.FirstOrDefault(item => item is ParseAsDateTimePipelineItem);
            if (dateTimeItem == null)
            {
                var item = new ParseAsDateTimePipelineItem().WithAdditionalFormats(formats);
                pipeline = pipeline.WithAdditionalItems(item);
            }
            else
            {
                dateTimeItem = dateTimeItem.WithAdditionalFormats(formats);
            }

            return pipeline;
        }

        public static TPipeline WithNewDelimiters<TPipeline, TEnumerable, TElement>(this TPipeline pipeline, params char[] delimiters) where TPipeline : SinglePipeline<TEnumerable> where TEnumerable : IEnumerable<TElement>
        {
            return pipeline.WithNewDelimiters<TPipeline, TEnumerable, TElement>((IEnumerable<char>)delimiters);
        }

        public static TPipeline WithNewDelimiters<TPipeline, TEnumerable, TElement>(this TPipeline pipeline, IEnumerable<char> delimiters) where TPipeline : SinglePipeline<TEnumerable> where TEnumerable : IEnumerable<TElement>
        {
            SplitWithDelimiterPipelineItem<TEnumerable, TElement> splitItem = (SplitWithDelimiterPipelineItem<TEnumerable, TElement>)pipeline.Items.FirstOrDefault(item => item is SplitWithDelimiterPipelineItem<TEnumerable, TElement>);
            if (splitItem == null)
            {
                var item = new SplitWithDelimiterPipelineItem<TEnumerable, TElement>(delimiters) as PipelineItem<TEnumerable>;
                pipeline = pipeline.WithAdditionalItems(item);
            }
            else
            {
                splitItem = splitItem.WithNewDelimiters(delimiters);
            }

            return pipeline;
        }
    }
}
