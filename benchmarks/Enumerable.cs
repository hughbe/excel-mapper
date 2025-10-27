using BenchmarkDotNet.Attributes;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

[MemoryDiagnoser]
public class Enumerable
{
    [Benchmark]
    public void ColumnIndices()
    {
        using var importer = Helpers.GetImporter("ManyColumns.xlsx");
        importer.Configuration.RegisterClassMap<ObjectArrayClass>(c =>
        {
            c.Map(p => p.Value)
                .WithColumnIndices(0, 1, 2, 3);
        });

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
