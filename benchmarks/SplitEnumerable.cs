using BenchmarkDotNet.Attributes;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

[MemoryDiagnoser]
public class SplitEnumerable
{
    private ExcelImporter _importer = default!;

    [IterationSetup]
    public void Setup()
    {
        _importer = Helpers.GetImporter("SplitWithManyCommas.xlsx");
    }
    
    [Benchmark]
    public void MapSplitEnumerable_AutoMapped()
    {
        var sheet = _importer.ReadSheet();
        sheet.ReadHeading();
        
        foreach (var value in sheet.ReadRows<ObjectArrayClass>())
        {
        }
    }

    [Benchmark]
    public void MapSplitEnumerable_StringSeparatorSingleChar()
    {
        _importer.Configuration.RegisterClassMap<ObjectArrayClass>(c =>
        {
            c.Map(p => p.Value)
                .WithSeparators(",");
        });

        var sheet = _importer.ReadSheet();
        sheet.ReadHeading();

        foreach (var value in sheet.ReadRows<ObjectArrayClass>())
        {
        }
    }

    [Benchmark]
    public void MapSplitEnumerable_StringSeparatorMultipleChars()
    {
        _importer.Configuration.RegisterClassMap<ObjectArrayClass>(c =>
        {
            c.Map(p => p.Value)
                .WithSeparators(",", ";");
        });

        var sheet = _importer.ReadSheet();
        sheet.ReadHeading();

        foreach (var value in sheet.ReadRows<ObjectArrayClass>())
        {
        }
    }

    [Benchmark]
    public void MapSplitEnumerable_CharSeparatorSingle()
    {
        _importer.Configuration.RegisterClassMap<ObjectArrayClass>(c =>
        {
            c.Map(p => p.Value)
                .WithSeparators(',');
        });

        var sheet = _importer.ReadSheet();
        sheet.ReadHeading();

        foreach (var value in sheet.ReadRows<ObjectArrayClass>())
        {
        }
    }

    [Benchmark]
    public void MapSplitEnumerable_CharSeparatorMultiple()
    {
        _importer.Configuration.RegisterClassMap<ObjectArrayClass>(c =>
        {
            c.Map(p => p.Value)
                .WithSeparators(',', ';');
        });

        var sheet = _importer.ReadSheet();
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
