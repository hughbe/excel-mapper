using System;

namespace ExcelMapper.Pipeline.Items
{
    public class ThrowIfStatusPipelineItem<T> : PipelineItem<T>
    {
        public PipelineStatus Status { get; }

        public ThrowIfStatusPipelineItem(PipelineStatus status)
        {
            if (!Enum.IsDefined(typeof(PipelineStatus), status))
            {
                throw new ArgumentException($"Invalid status type {status}.", nameof(status));
            }

            Status = status;
        }

        public override PipelineResult<T> TryMap(PipelineResult<T> item)
        {
            if (item.Status == Status)
            {
                throw new ExcelMappingException($"{item.Status} result for parameter \"{item.Context.StringValue}\" of type \"{typeof(T)}\"", item.Context);
            }

            return item;
        }
    }
}
