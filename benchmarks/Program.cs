using BenchmarkDotNet.Running;

namespace ExcelMapper.Benchmarks;

public class Program
{
    // Run a subset with, for example: dotnet run -c Release -- --filter *Mapping*
    public static void Main(string[] args) => BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
}
