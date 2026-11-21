using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Tests;

public class OneToOneMapTTests
{





    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Optional_Set_GetReturnsExpected(bool value)
    {
        var map = CreateMap();
        map.Optional = value;
        Assert.Equal(value, map.Optional);

        // Set same.
        map.Optional = value;
        Assert.Equal(value, map.Optional);

        // Set different.
        map.Optional = !value;
        Assert.Equal(!value, map.Optional);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PreserveFormatting_Set_GetReturnsExpected(bool value)
    {
        var map = CreateMap();
        map.PreserveFormatting = value;
        Assert.Equal(value, map.PreserveFormatting);

        // Set same.
        map.PreserveFormatting = value;
        Assert.Equal(value, map.PreserveFormatting);

        // Set different.
        map.PreserveFormatting = !value;
        Assert.Equal(!value, map.PreserveFormatting);
    }



    [Fact]
    public void TryGetValue_InvokeCantReadPropertyInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        sheet.ReadHeading();
        var map = CreateMap().WithColumnName("NoSuchColumn");
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadFieldInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        sheet.ReadHeading();
        var map = CreateMap().WithColumnName("NoSuchColumn");
        MemberInfo member = typeof(TestClass).GetField(nameof(TestClass._field))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadEventInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        sheet.ReadHeading();
        var map = CreateMap().WithColumnName("NoSuchColumn");
        MemberInfo member = typeof(TestClass).GetEvent(nameof(TestClass.Event))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMember_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        sheet.ReadHeading();
        var map = CreateMap().WithColumnName("NoSuchColumn");
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberColumnName_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var map = CreateMap().WithColumnName("NoSuchColumn");
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberColumnIndex_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Reader.Read();

        var map = CreateMap().WithColumnIndex(int.MaxValue);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    private static OneToOneMap<string> CreateMap() => new ExcelClassMap<TestClass>().Map(t => t.Value);

    [Fact]
    public void ColumnMethods_Invoke_ReadExpectedColumn()
    {
        using var importer = Helpers.GetCsvImporter("Name,Other Name,Year 2024\na,b,c\n");
        string Read(Func<OneToOneMap<string>, OneToOneMap<string>> configure)
        {
            var map = new ExcelClassMap<TestClass>();
            Assert.Same(configure(map.Map(t => t.Value)), map.Properties.Single().Map);
            using var reader = Helpers.GetCsvImporter("Name,Other Name,Year 2024\na,b,c\n");
            reader.Configuration.RegisterClassMap(map);
            return reader.ReadSheet().ReadRows<TestClass>().Single().Value;
        }

        Assert.Equal("a", Read(m => m.WithColumnName("NAME")));
        Assert.Equal("a", Read(m => m.WithColumnName("Name", StringComparison.Ordinal)));
        Assert.Equal("b", Read(m => m.WithColumnNameMatching("Missing", "Other Name")));
        Assert.Equal("b", Read(m => m.WithColumnNameMatching(n => n.StartsWith("Other"))));
        Assert.Equal("c", Read(m => m.WithColumnNameMatching(new System.Text.RegularExpressions.Regex(@"Year \d+"))));
        Assert.Equal("b", Read(m => m.WithColumnIndex(1)));
        Assert.Equal("c", Read(m => m.WithColumnLetter("C")));
        Assert.Equal("c", Read(m => m.WithColumns(new LastColumnSelector())));
    }

    [Fact]
    public void ColumnMethods_InvalidArguments_Throw()
    {
        var map = CreateMap();
        Assert.Throws<ArgumentNullException>("columnName", () => map.WithColumnName(null!));
        Assert.Throws<ArgumentException>("columnName", () => map.WithColumnName(""));
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => map.WithColumnName("Name", (StringComparison)99));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNameMatching((string[])null!));
        Assert.Throws<ArgumentNullException>("predicate", () => map.WithColumnNameMatching((Func<string, bool>)null!));
        Assert.Throws<ArgumentNullException>("regex", () => map.WithColumnNameMatching((System.Text.RegularExpressions.Regex)null!));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => map.WithColumnIndex(-1));
        Assert.Throws<ArgumentNullException>("columnLetter", () => map.WithColumnLetter(null!));
        Assert.Throws<ArgumentException>("columnLetter", () => map.WithColumnLetter("1"));
        Assert.Throws<ArgumentNullException>("selector", () => map.WithColumns(null!));
    }

    [Fact]
    public void MakeOptional_MakePreserveFormatting_SetProperties()
    {
        var map = CreateMap();
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
    }

    private sealed class LastColumnSelector : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ExcelSheet sheet) => [sheet.NumberOfColumns - 1];

        public string Describe(ExcelSheet sheet) => "the last column";
    }

    private class TestClass
    {
        public string Value { get; set; } = default!;
#pragma warning disable 0649
        public string _field = default!;
#pragma warning restore 0649

        public event EventHandler Event { add { } remove { } }
    }
}
