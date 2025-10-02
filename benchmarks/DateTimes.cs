using System;
using BenchmarkDotNet.Attributes;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

[MemoryDiagnoser]
public class DateTimes
{
    [Benchmark]
    public void DefaultMap()
    {
        using var original = Helpers.GetResource("ManyDates.xlsx");
        var importer = new ExcelImporter(original);

        ExcelSheet sheet = importer.ReadSheet();
        foreach (object value in sheet.ReadRows<DataClass>())
        {
        }
    }

    private class DataClass
    {
        public DateTime Value { get; set; }
    }
}
