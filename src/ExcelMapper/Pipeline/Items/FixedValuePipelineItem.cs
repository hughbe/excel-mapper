namespace ExcelMapper.Pipeline.Items
{
    public class FixedValuePipelineItem<T> : PipelineItem<T>
    {
        public T Value { get; }

        public FixedValuePipelineItem(T value) => Value = value;

        public override PipelineResult<T> TryMap(PipelineResult<T> item)
        {
            return item.MakeCompleted(Value);
        }
    }
}
