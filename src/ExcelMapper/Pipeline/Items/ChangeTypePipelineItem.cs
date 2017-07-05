using System;

namespace ExcelMapper.Pipeline.Items
{
    public class ChangeTypePipelineItem<T> : PipelineItem<T>
    {
        public Type Type { get; }

        public ChangeTypePipelineItem(Type type)
        {
            Type = type;
        }

        public override PipelineResult<T> TryMap(PipelineResult<T> item)
        {
            if (string.IsNullOrEmpty(item.Context.StringValue))
            {
                return item.MakeEmpty();
            }

            try
            {
                T value = (T)Convert.ChangeType(item.Context.StringValue, Type);
                return item.MakeCompleted(value);
            }
            catch
            {
                return item.MakeInvalid();
            }
        }
    }
}
