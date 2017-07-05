using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelMapper.Pipeline.Items
{
    public class ParseAsDateTimePipelineItem : PipelineItem<DateTime>
    {
        /// <summary>
        /// Defaults to "G" - the default Excel format.
        /// </summary>
        public string[] Formats { get; private set; } = new string[] { "G" };
        public IFormatProvider Provider { get; private set; }
        public DateTimeStyles Style { get; private set; }

        public ParseAsDateTimePipelineItem WithAdditionalFormats(params string[] formats) => WithAdditionalFormats((IEnumerable<string>)formats);

        public ParseAsDateTimePipelineItem WithAdditionalFormats(IEnumerable<string> formats)
        {
            if (formats == null)
            {
                throw new ArgumentNullException(nameof(formats));
            }

            Formats = Formats.Concat(formats).ToArray();
            return this;
        }

        public ParseAsDateTimePipelineItem WithProvider(IFormatProvider provider)
        {
            Provider = provider;
            return this;
        }

        public ParseAsDateTimePipelineItem WithStyle(DateTimeStyles style)
        {
            Style = style;
            return this;
        }

        public override PipelineResult<DateTime> TryMap(PipelineResult<DateTime> item)
        {
            if (string.IsNullOrEmpty(item.Context.StringValue))
            {
                return item.MakeEmpty();
            }

            if (!DateTime.TryParseExact(item.Context.StringValue, Formats, Provider, Style, out DateTime result))
            {
                return item.MakeInvalid();
            }

            return item.MakeCompleted(result);
        }
    }
}
