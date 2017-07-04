namespace ExcelMapper.Pipeline.Items
{
    public class ParseAsBoolPipelineItem : PipelineItem<bool>
    {
        public override PipelineResult<bool> TryMap(PipelineResult<bool> item)
        {
            if (string.IsNullOrEmpty(item.StringValue))
            {
                return item.MakeEmpty();
            }

            if (!bool.TryParse(item.StringValue, out bool result))
            {
                if (item.StringValue == "1")
                {
                    return item.MakeCompleted(true);
                }
                else if (item.StringValue == "0")
                {
                    return item.MakeCompleted(false);
                }

                return item.MakeInvalid();
            }

            return item.MakeCompleted(result);
        }
    }
}
