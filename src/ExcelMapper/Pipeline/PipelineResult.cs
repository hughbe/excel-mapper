namespace ExcelMapper.Pipeline
{
    public struct PipelineResult<T>
    {
        public PipelineStatus Status { get; }
        public PipelineContext Context { get; }
        public T Result { get; }

        public PipelineResult(PipelineStatus status, PipelineContext context, T result)
        {
            Status = status;
            Context = context;
            Result = result;
        }

        public PipelineResult<T> MakeEmpty() => new PipelineResult<T>(PipelineStatus.Empty, Context, Result);

        public PipelineResult<T> MakeInvalid() => new PipelineResult<T>(PipelineStatus.Invalid, Context, Result);

        public PipelineResult<T> MakeSuccess(PipelineContext context) => new PipelineResult<T>(PipelineStatus.Success, context, Result);

        public PipelineResult<T> MakeCompleted(T result) => new PipelineResult<T>(PipelineStatus.Completed, Context, result);
    }
}
