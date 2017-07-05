namespace ExcelMapper.Pipeline
{
    public class TrimStringValuePipelineItem<T> : PipelineItem<T>
    {
        public override PipelineResult<T> TryMap(PipelineResult<T> item)
        {
            if (string.IsNullOrEmpty(item.Context.StringValue))
            {
                return item.MakeEmpty();
            }

            item.Context.StringValue = item.Context.StringValue.Trim();
            return item.MakeSuccess(item.Context);
        }
    }
}
