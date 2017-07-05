using System;
using System.Globalization;

namespace ExcelMapper.Pipeline.Items
{
    public class ParseAsIntPipelineItem : PipelineItem<int>
    {
        public NumberStyles Style { get; private set; }
        public IFormatProvider Provider { get; private set; }

        public ParseAsIntPipelineItem WithStyle(NumberStyles style)
        {
            Style = style;
            return this;
        }

        public ParseAsIntPipelineItem WithProvider(IFormatProvider provider)
        {
            Provider = provider;
            return this;
        }

        public override PipelineResult<int> TryMap(PipelineResult<int> item)
        {
            if (string.IsNullOrEmpty(item.Context.StringValue))
            {
                return item.MakeEmpty();
            }

            if (!int.TryParse(item.Context.StringValue, out int result))
            {
                return item.MakeInvalid();
            }

            return item.MakeCompleted(result);
        }
    }
}
