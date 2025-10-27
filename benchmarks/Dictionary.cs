using System.Collections.Generic;
using BenchmarkDotNet.Attributes;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

[MemoryDiagnoser]
public class Dictionary
{
    private ExcelImporter _importer = null!;

    [IterationSetup]
    public void Setup()
    {
        _importer = Helpers.GetImporter("ManyColumns.xlsx");
    }

    [Benchmark]
    public void DictionaryStringObject()
    {
        var sheet = _importer.ReadSheet();
        sheet.ReadHeading();

        foreach (var value in sheet.ReadRows<Dictionary<string, object>>())
        {
        }
    }
}
