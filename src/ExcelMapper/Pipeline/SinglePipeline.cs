using System.Reflection;

namespace ExcelMapper.Pipeline
{
    public abstract class SinglePipeline<T> : Pipeline<T>
    {
        public SinglePipeline(MemberInfo member) : base(member) { }
    }
}
