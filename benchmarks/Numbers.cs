using BenchmarkDotNet.Attributes;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

[MemoryDiagnoser]
public class Numbers
{
    [Benchmark]
    public void DefaultMap()
    {
        using var original = Helpers.GetResource("ManyNumbers.xlsx");
        var importer = new ExcelImporter(original);

        ExcelSheet sheet = importer.ReadSheet();
        foreach (object value in sheet.ReadRows<DataClass>())
        {
        }
    }

    private class DataClass
    {
        public int Value { get; set; }
    }
}
