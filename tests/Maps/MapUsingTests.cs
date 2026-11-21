using System.Linq;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Tests;

public class MapUsingTests
{
    [Fact]
    public void ReadRow_ConvertUsingMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<ConvertUsingValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConvertUsingValue>();
        Assert.Equal("aextra", row1.StringValue);
    }

    private class ConvertUsingValue
    {
        public string StringValue { get; set; } = default!;
    }

    private class ConvertUsingValueMap : ExcelClassMap<ConvertUsingValue>
    {
        public ConvertUsingValueMap()
        {
            Map(c => c.StringValue)
                .WithConverter(s => s + "extra");
        }
    }

    // A converter replaces the built-in conversion for every type, so it's used for values the
    // built-in conversion could also convert.
    [Fact]
    public void ReadRows_ConverterOnBuiltInTypes_UsedForEveryValue()
    {
        using var importer = Helpers.GetCsvImporter("Number,Nullable,When,Day,Name\n3,4,2024-01-02,Monday,a\n");
        importer.Configuration.RegisterClassMap<BuiltInTypes>(c =>
        {
            c.Map(r => r.Number).WithConverter(s => int.Parse(s!) * 2);
            c.Map(r => r.Nullable).WithConverter(s => (int?)int.Parse(s!) * 10);
            c.Map(r => r.When).WithConverter(s => DateTime.Parse(s!).AddDays(1));
            c.Map(r => r.Day).WithConverter(s => s == "Monday" ? DayOfWeek.Friday : DayOfWeek.Sunday);
            c.Map(r => r.Name).WithConverter(s => s!.ToUpperInvariant());
        });
        var row = importer.ReadSheet().ReadRows<BuiltInTypes>().Single();
        Assert.Equal(6, row.Number);
        Assert.Equal(40, row.Nullable);
        Assert.Equal(new DateTime(2024, 1, 3), row.When);
        Assert.Equal(DayOfWeek.Friday, row.Day);
        Assert.Equal("A", row.Name);
    }

    [Fact]
    public void ReadRows_ConverterReplacesAttributeMapper()
    {
        using var importer = Helpers.GetCsvImporter("Number\n3\n");
        importer.Configuration.RegisterClassMap<AttributeMapped>(c => c.Map(r => r.Number).WithConverter(s => int.Parse(s!) + 100));
        Assert.Equal(103, importer.ReadSheet().ReadRows<AttributeMapped>().Single().Number);
    }

    [Fact]
    public void ReadRows_ConverterOnEmptyCell_UsesEmptyFallback()
    {
        using var importer = Helpers.GetCsvImporter("Number,Nullable,When,Day,Name\n,,2024-01-02,Monday,a\n");
        importer.Configuration.RegisterClassMap<BuiltInTypes>(c =>
        {
            c.Map(r => r.Number).WithConverter(_ => throw new InvalidOperationException("Not called for empty cells.")).WithEmptyFallback(-1);
            c.Map(r => r.Nullable).WithConverter(_ => throw new InvalidOperationException("Not called for empty cells."));
            c.Map(r => r.When);
            c.Map(r => r.Day);
            c.Map(r => r.Name);
        });
        var row = importer.ReadSheet().ReadRows<BuiltInTypes>().Single();
        Assert.Equal(-1, row.Number);
        Assert.Null(row.Nullable);
    }

    [Fact]
    public void ReadRows_ConverterThrows_UsesInvalidFallback()
    {
        using var importer = Helpers.GetCsvImporter("Number,Nullable,When,Day,Name\n3,x,2024-01-02,Monday,a\n");
        importer.Configuration.RegisterClassMap<BuiltInTypes>(c =>
        {
            c.Map(r => r.Number).WithConverter(_ => throw new FormatException()).WithInvalidFallback(-1);
            c.Map(r => r.Nullable).WithConverter((string? s) => (int?)int.Parse(s!));
            c.Map(r => r.When);
            c.Map(r => r.Day);
            c.Map(r => r.Name);
        });
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<BuiltInTypes>().ToArray());
        Assert.Equal("B2", exception.CellAddress);
        Assert.IsType<FormatException>(exception.InnerException);
    }

    [Fact]
    public void ReadRows_ConverterOnElementsAndDictionaryValues_UsedForEveryValue()
    {
        using var importer = Helpers.GetCsvImporter("Values,A,B\n1;2,3,4\n");
        importer.Configuration.RegisterClassMap<Collections>(c =>
        {
            c.Map(r => r.Values).WithSeparators(';').WithElementMap(e => e.WithConverter(s => int.Parse(s!) * 2));
            c.Map(r => r.Columns).WithColumnNames("A", "B").WithValueMap(v => v.WithConverter(s => int.Parse(s!) * 10));
        });
        var row = importer.ReadSheet().ReadRows<Collections>().Single();
        Assert.Equal([2, 4], row.Values);
        Assert.Equal(new Dictionary<string, int> { ["A"] = 30, ["B"] = 40 }, row.Columns);
    }

    public class BuiltInTypes
    {
        public int Number { get; set; }
        public int? Nullable { get; set; }
        public DateTime When { get; set; }
        public DayOfWeek Day { get; set; }
        public string? Name { get; set; }
    }

    public class AttributeMapped
    {
        [ExcelMapper(typeof(TimesThreeMapper))]
        public int Number { get; set; }
    }

    public class TimesThreeMapper : ICellMapper
    {
        public CellMapperResult Map(ReadCellResult readResult) => CellMapperResult.Success(int.Parse(readResult.GetString()!) * 3);
    }

    public class Collections
    {
        public int[] Values { get; set; } = [];
        public Dictionary<string, int> Columns { get; set; } = [];
    }
}
