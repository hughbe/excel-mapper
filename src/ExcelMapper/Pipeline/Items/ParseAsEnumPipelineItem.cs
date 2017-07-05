using System;

namespace ExcelMapper.Pipeline.Items
{
    public class ParseAsEnumPipelineItem<TEnum> : PipelineItem<TEnum> where TEnum : struct
    {
        public override PipelineResult<TEnum> TryMap(PipelineResult<TEnum> item)
        {
            if (string.IsNullOrEmpty(item.Context.StringValue))
            {
                return item.MakeEmpty();
            }

            if (!Enum.TryParse(item.Context.StringValue, out TEnum result))
            {
                return item.MakeInvalid();
            }

            return item.MakeCompleted(result);
        }
    }
}
