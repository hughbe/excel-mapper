namespace ExcelMapper.Pipeline.Items
{
    public class ParseAsBoolPipelineItem : PipelineItem<bool>
    {
        public override PipelineResult<bool> TryMap(PipelineResult<bool> item)
        {
            if (string.IsNullOrEmpty(item.Context.StringValue))
            {
                return item.MakeEmpty();
            }

            // Excel transforms bool values such as "true" or "false" to "1" or "0".
            if (item.Context.StringValue == "1")
            {
                return item.MakeCompleted(true);
            }
            else if (item.Context.StringValue == "0")
            {
                return item.MakeCompleted(false);
            }

            if (!bool.TryParse(item.Context.StringValue, out bool result))
            {
                return item.MakeInvalid();
            }

            return item.MakeCompleted(result);
        }
    }
}
