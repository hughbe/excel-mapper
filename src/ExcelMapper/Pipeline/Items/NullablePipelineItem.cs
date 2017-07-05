using System.Reflection;

namespace ExcelMapper.Pipeline.Items
{
    internal class CastPipelineItem<T, U> : PipelineItem<T>
    {
        public PipelineItem<U> Item { get; }

        public CastPipelineItem(PipelineItem<U> item)
        {
            Item = item;
        }

        public override PipelineResult<T> TryMap(PipelineResult<T> item)
        {
            U value;
            if (item.Result == null)
            {
                value = default(U);
            }
            else
            {
                value = (U)(object)item.Result;
            }

            PipelineResult<U> result = new PipelineResult<U>(item.Status, item.Context,value);
            result = Item.TryMap(result);

            if (result.Status == PipelineStatus.Completed)
            {
                return new PipelineResult<T>(result.Status, result.Context, (T)(object)result.Result);
            }

            return new PipelineResult<T>(result.Status, result.Context, item.Result);
        }
    }
}
