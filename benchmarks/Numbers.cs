using BenchmarkDotNet.Attributes;
using ExcelDataReader;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

[MemoryDiagnoser]
public class Numbers
{
    [Benchmark]
    public void ExcelDataReaderMap_GetDouble()
    {
        using var original = Helpers.GetResource("ManyNumbers.xlsx");
        var reader = ExcelReaderFactory.CreateReader(original);
        var importer = new ExcelImporter(reader);
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Read the numbers.
        while (reader.Read())
        {
            _ = reader.GetDouble(0);
        }
    }

    [Benchmark]
    public void ExcelDataReaderMap_GetObject()
    {
        using var original = Helpers.GetResource("ManyNumbers.xlsx");
        var reader = ExcelReaderFactory.CreateReader(original);
        var importer = new ExcelImporter(reader);
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Read the numbers.
        while (reader.Read())
        {
            _ = reader.GetValue(0);
        }
    }

    [Benchmark]
    public void DefaultMap()
    {
        using var original = Helpers.GetResource("ManyNumbers.xlsx");
        var importer = new ExcelImporter(original);

        var sheet = importer.ReadSheet();
        foreach (object value in sheet.ReadRows<DataClass>())
        {
        }
    }

    private class DataClass
    {
        public int Value { get; set; }
    }
}
