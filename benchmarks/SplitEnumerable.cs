using BenchmarkDotNet.Attributes;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

[MemoryDiagnoser]
public class SplitEnumerable
{
    [Benchmark]
    public void ExcelDataReaderMap_GetDouble()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        foreach (var value in sheet.ReadRows<ObjectArrayClass>())
        {
        }
    }

    private class ObjectArrayClass
    {
        public string[] Value { get; set; } = default!;
    }
}
