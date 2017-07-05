namespace ExcelMapper.Pipeline.Items
{
    public class ParseAsStringPipelineItem : PipelineItem<string>
    {
        private ParseAsStringPipelineItem() { }

        private static ParseAsStringPipelineItem s_instance = null;
        public static ParseAsStringPipelineItem Instance => s_instance ?? (s_instance = new ParseAsStringPipelineItem());

        public override PipelineResult<string> TryMap(PipelineResult<string> item)
        {
            return item.MakeCompleted(item.Context.StringValue);
        }
    }
}
