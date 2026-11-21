using System;
using System.Globalization;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

/// <summary>
/// Measures the cost of ExcelImporterConfiguration.Culture, which sets the culture for each row.
/// </summary>
[MemoryDiagnoser]
public class CultureMapping
{
    private readonly Mapping _mapping = new();
    private byte[] _csv = default!;

    [GlobalSetup]
    public void Setup()
    {
        _mapping.Setup();
        _csv = (byte[])typeof(Mapping).GetField("_validCsv", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!.GetValue(_mapping)!;

        // Make sure the configured culture differs from the current one, so it is switched every row.
        CultureInfo.CurrentCulture = new CultureInfo("en-GB");
    }

    [Benchmark(Baseline = true)]
    public int CultureNotSet() => Read(null);

    [Benchmark]
    public int CultureSet() => Read(CultureInfo.InvariantCulture);

    private int Read(CultureInfo? culture)
    {
        using var importer = new ExcelImporter(new MemoryStream(_csv), ExcelImporterFileType.Csv);
        importer.Configuration.Culture = culture;
        importer.Configuration.RegisterClassMap<Mapping.Employee>(c =>
        {
            c.Map(p => p.Name);
            c.Map(p => p.Age);
            c.Map(p => p.Score);
            c.Map(p => p.Salary);
            c.Map(p => p.Active);
            c.Map(p => p.Id);
            c.Map(p => p.Status);
            c.Map(p => p.Manager);
            c.Map(p => p.Notes);
        });

        return importer.ReadSheet().ReadRows<Mapping.Employee>().Sum(r => r.Age);
    }
}
