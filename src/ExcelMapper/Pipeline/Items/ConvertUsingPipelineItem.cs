using System;

namespace ExcelMapper.Pipeline.Items
{
    public delegate T ConvertUsingSimple<T>(string stringValue);

    public class ConvertUsingPipelineItem<T> : PipelineItem<T>
    {
        public Func<PipelineResult<T>, PipelineResult<T>> Converter { get; }

        public ConvertUsingPipelineItem(Func<PipelineResult<T>, PipelineResult<T>> converter)
        {
            Converter = converter ?? throw new ArgumentNullException(nameof(converter));
        }

        public ConvertUsingPipelineItem(ConvertUsingSimple<T> converter)
        {
            if (converter == null)
            {
                throw new ArgumentNullException(nameof(converter));
            }

            Converter = item =>
            {
                T value = converter(item.Context.StringValue);
                return item.MakeCompleted(value);
            };
        }

        public override PipelineResult<T> TryMap(PipelineResult<T> item) => Converter(item);
    }
}
