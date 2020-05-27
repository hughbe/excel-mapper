using BenchmarkDotNet.Running;

namespace ExcelMapper.Benchmarks
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<ReadRows>();
        }
    }
}
