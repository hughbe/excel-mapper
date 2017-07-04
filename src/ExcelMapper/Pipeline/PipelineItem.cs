namespace ExcelMapper.Pipeline
{
    public abstract class PipelineItem<T>
    {
        internal bool Automapped { get; set; }

        public abstract PipelineResult<T> TryMap(PipelineResult<T> item);
    }
}
