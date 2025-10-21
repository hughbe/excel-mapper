using BenchmarkDotNet.Attributes;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

[MemoryDiagnoser]
public class ReadRows
{
    [Benchmark]
    public void DefaultMap()
    {
        using var original = Helpers.GetResource("VeryLargeSheet.xlsx");
        var importer = new ExcelImporter(original);
        importer.Configuration.RegisterClassMap<DataClassMap>();

        var sheet = importer.ReadSheet();
        foreach (object value in sheet.ReadRows<DataClass>())
        {
        }
    }

    [Benchmark]
    public void SkipBlankLinesMap()
    {
        using var original = Helpers.GetResource("VeryLargeSheet.xlsx");
        var importer = new ExcelImporter(original);
        importer.Configuration.SkipBlankLines = true;
        importer.Configuration.RegisterClassMap<DataClassMap>();

        var sheet = importer.ReadSheet();
        foreach (object value in sheet.ReadRows<DataClass>())
        {
        }
    }

    [Benchmark]
    public void OptionalMap()
    {
        using var original = Helpers.GetResource("VeryLargeSheet.xlsx");
        var importer = new ExcelImporter(original);
        importer.Configuration.RegisterClassMap<OptionalDataClassMap>();

        var sheet = importer.ReadSheet();
        foreach (object value in sheet.ReadRows<DataClass>())
        {
        }
    }

    [Benchmark]
    public void OptionalNoSuchValueMap()
    {
        using var original = Helpers.GetResource("VeryLargeSheet.xlsx");
        var importer = new ExcelImporter(original);
        importer.Configuration.RegisterClassMap<OptionalDataWithMissingValueClassMap>();

        var sheet = importer.ReadSheet();
        foreach (object value in sheet.ReadRows<DataClassWithMissingValue>())
        {
        }
    }

    private class DataClass
    {
        public int Value { get; set; }
    }

    private class DataClassWithMissingValue
    {
        public int Value { get; set; }
        public int NoSuchValue { get; set; }
    }

    private class DataClassMap : ExcelClassMap<DataClass>
    {
        public DataClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class OptionalDataClassMap : ExcelClassMap<DataClass>
    {
        public OptionalDataClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class OptionalDataWithMissingValueClassMap : ExcelClassMap<DataClassWithMissingValue>
    {
        public OptionalDataWithMissingValueClassMap()
        {
            Map(p => p.Value);

            Map(p => p.NoSuchValue)
                .MakeOptional()
                .WithEmptyFallback(0);
        }
    }
}
