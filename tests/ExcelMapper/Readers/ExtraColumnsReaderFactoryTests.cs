using ExcelMapper.Abstractions;
using ExcelMapper.Tests;

namespace ExcelMapper.Readers.Tests;

public class ExtraColumnsReaderFactoryTests
{
    [Fact]
    public void Ctor_ClassMap()
    {
        var classMap = new ExcelClassMap<TestClass>();
        var factory = new ExtraColumnsReaderFactory(classMap);
        Assert.Same(classMap, factory.ClassMap);
    }

    [Fact]
    public void Ctor_NullClassMap_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("classMap", () => new ExtraColumnsReaderFactory(null!));
    }

    [Fact]
    public void GetCellsReader_InvokeNoOtherMembers_ReturnsAllColumns()
    {
        using var importer = Helpers.GetCsvImporter("A,B,,C\n1,2,3,4\n");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        importer.Reader.Read();

        var factory = new ExtraColumnsReaderFactory(new ExcelClassMap<TestClass>());
        var reader = factory.GetCellsReader(sheet);
        Assert.NotNull(reader);
        for (var i = 0; i < 2; i++)
        {
            Assert.True(reader.Start(importer.Reader, false, out var count));
            Assert.Equal(3, count);
            Assert.Equal(["1", "2", "4"], ReadAll(reader));
            reader.Reset();
        }
    }

    [Fact]
    public void GetCellsReader_InvokeOtherMembers_ReturnsUnreadColumns()
    {
        using var importer = Helpers.GetCsvImporter("A,B,C,D,E,F\n1,2,3,4,5,6\n");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        importer.Reader.Read();

        var classMap = new ExcelClassMap<TestClass>();
        classMap.Map(p => p.Value).WithColumnName("b");
        classMap.Map(p => p.Values).WithColumnNames("D", "C");
        classMap.Map(p => p.Optional).WithColumnName("Missing").MakeOptional();
        classMap.MapObject(p => p.Child).Map(c => c.Value).WithColumnIndex(5);
        var extraMap = classMap.MapExtraColumns(p => p.Extra);

        var reader = Assert.IsAssignableFrom<ICellsReader>(extraMap.ReaderFactory.GetCellsReader(sheet));
        Assert.True(reader.Start(importer.Reader, false, out var count));
        Assert.Equal(2, count);
        Assert.Equal(["1", "5"], ReadAll(reader));
    }

    [Fact]
    public void GetCellsReader_InvokeEveryColumnRead_ReturnsEmptyReader()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        importer.Reader.Read();

        var classMap = new ExcelClassMap<TestClass>();
        classMap.Map(p => p.Value);
        var factory = new ExtraColumnsReaderFactory(classMap);
        var reader = factory.GetCellsReader(sheet);
        Assert.NotNull(reader);
        Assert.True(reader.Start(importer.Reader, false, out var count));
        Assert.Equal(0, count);
        Assert.False(reader.TryGetNext(out _));
        reader.Reset();
    }

    [Fact]
    public void GetCellsReader_InvokeNullSheet_ThrowsArgumentNullException()
    {
        var factory = new ExtraColumnsReaderFactory(new ExcelClassMap<TestClass>());
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetCellsReader_InvokeSheetWithoutHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        var factory = new ExtraColumnsReaderFactory(new ExcelClassMap<TestClass>());
        Assert.Throws<ExcelMappingException>(() => factory.GetCellsReader(sheet));
    }

    private static List<string?> ReadAll(ICellsReader reader)
    {
        var values = new List<string?>();
        while (reader.TryGetNext(out var result))
        {
            values.Add(result.StringValue);
        }

        return values;
    }

    private class TestClass
    {
        public string? Value { get; set; }
        public string[]? Values { get; set; }
        public string? Optional { get; set; }
        public ChildClass? Child { get; set; }
        public Dictionary<string, string>? Extra { get; set; }
    }

    private class ChildClass
    {
        public string? Value { get; set; }
    }
}
