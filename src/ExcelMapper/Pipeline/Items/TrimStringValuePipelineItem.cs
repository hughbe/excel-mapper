namespace ExcelMapper.Pipeline
{
    public class TrimStringValuePipelineItem<T> : PipelineItem<T>
    {
        public override PipelineResult<T> TryMap(PipelineResult<T> item)
        {
            if (string.IsNullOrEmpty(item.StringValue))
            {
                return item.MakeEmpty();
            }

            return item.MakeSuccess(item.StringValue.Trim());
        }
    }
}
