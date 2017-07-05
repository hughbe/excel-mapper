namespace ExcelMapper.Pipeline.Items
{
    public class ParseAsStringPipelineItem : PipelineItem<string>
    {
        private ParseAsStringPipelineItem() { }

        private static ParseAsStringPipelineItem s_instance = null;
        public static ParseAsStringPipelineItem Instance => s_instance ?? (s_instance = new ParseAsStringPipelineItem());

        public override PipelineResult<string> TryMap(PipelineResult<string> item)
        {
            // Set the Result value to Context.StringValue.
            return new PipelineResult<string>(PipelineStatus.Success, item.Context, item.Context.StringValue);
        }
    }
}
