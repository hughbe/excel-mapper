using System.Linq;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Tests;

public class MapCustomColumnMatchingTests
{
    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNamesEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerablePropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Year 2023", "Year 2024"}])]
        public object?[] CustomName { get; set; } = default!;
    }

    private class CustomEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("Year 2023", "Year 2024"));
        }
    }

    private class DefaultCustomNamesEnumerablePropertyClassMap : ExcelClassMap<CustomNamesEnumerablePropertyClass>
    {
        public DefaultCustomNamesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class EnumerablePropertyClass
    {
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNamesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class EnumerableFieldClass
    {
        public object?[] CustomName = default!;
    }

    private class CustomNamesEnumerableFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Year 2023", "Year 2024"}])]
        public object?[] CustomName = default!;
    }

    private class DefaultCustomNamesEnumerableFieldClassMap : ExcelClassMap<CustomNamesEnumerableFieldClass>
    {
        public DefaultCustomNamesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("Year 2023", "Year 2024"));
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNamesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingCustomNamesEnumerablePropertyClass>();
        Assert.Equal(["1"], row1.CustomName);   
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNamesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomNamesEnumerablePropertyClassMap>();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingCustomNamesEnumerablePropertyClass>();
        Assert.Equal(["1"], row1.CustomName);   
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<CustomNoMatchingEnumerablePropertyClassMap>();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1"], row1.CustomName);
    }

    private class NoMatchingCustomNamesEnumerablePropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Year 2023", "NoSuchColumn"}])]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultNoMatchingCustomNamesEnumerablePropertyClassMap : ExcelClassMap<NoMatchingCustomNamesEnumerablePropertyClass>
    {
        public DefaultNoMatchingCustomNamesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoMatchingEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomNoMatchingEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("Year 2023", "NoSuchColumn"));
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingCustomNamesEnumerableFieldClass>();
        Assert.Equal(["1"], row1.CustomName);   
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomNamesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingCustomNamesEnumerableFieldClass>();
        Assert.Equal(["1"], row1.CustomName);   
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1"], row1.CustomName);
    }

    private class NoMatchingCustomNamesEnumerableFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Year 2023", "NoSuchColumn"}])]
        public object?[] CustomName = default!;
    }

    private class DefaultNoMatchingCustomNamesEnumerableFieldClassMap : ExcelClassMap<NoMatchingCustomNamesEnumerableFieldClass>
    {
        public DefaultNoMatchingCustomNamesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoMatchingEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomNoMatchingEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("Year 2023", "NoSuchColumn"));
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomNamesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesEnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomNamesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingCustomNamesEnumerablePropertyClassMap>();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesEnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoneMatchingEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerablePropertyClass>());
    }

    private class NoneMatchingCustomNamesEnumerablePropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "NoSuchColumn", "NoSuchColumn"}])]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultNoneMatchingCustomNamesEnumerablePropertyClassMap : ExcelClassMap<NoneMatchingCustomNamesEnumerablePropertyClass>
    {
        public DefaultNoneMatchingCustomNamesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoneMatchingEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomNoneMatchingEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("NoSuchColumn", "NoSuchColumn"));
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomNamesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingCustomNamesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoneMatchingEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerableFieldClass>());
    }

    private class NoneMatchingCustomNamesEnumerableFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "NoSuchColumn", "NoSuchColumn"}])]
        public object?[] CustomName = default!;
    }

    private class DefaultNoneMatchingCustomNamesEnumerableFieldClassMap : ExcelClassMap<NoneMatchingCustomNamesEnumerableFieldClass>
    {
        public DefaultNoneMatchingCustomNamesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoneMatchingEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomNoneMatchingEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("NoSuchColumn", "NoSuchColumn"));
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNamesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesEnumerablePropertyClass>();
        Assert.Equal(["1"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNamesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomNamesEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesEnumerablePropertyClass>();
        Assert.Equal(["1"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingOptionalEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1"], row1.CustomName);
    }

    private class NoMatchingOptionalCustomNamesEnumerablePropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Year 2023", "NoSuchColumn"}])]
        [ExcelOptional]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultNoMatchingOptionalCustomNamesEnumerablePropertyClassMap : ExcelClassMap<NoMatchingOptionalCustomNamesEnumerablePropertyClass>
    {
        public DefaultNoMatchingOptionalCustomNamesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoMatchingOptionalEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomNoMatchingOptionalEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("Year 2023", "NoSuchColumn"))
                .MakeOptional();
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesEnumerableFieldClass>();
        Assert.Equal(["1"], row1.CustomName);
    }
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomNamesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesEnumerableFieldClass>();
        Assert.Equal(["1"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingOptionalEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1"], row1.CustomName);
    }

    private class NoMatchingOptionalCustomNamesEnumerableFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Year 2023", "NoSuchColumn"}])]
        [ExcelOptional]
        public object?[] CustomName = default!;
    }

    private class DefaultNoMatchingOptionalCustomNamesEnumerableFieldClassMap : ExcelClassMap<NoMatchingOptionalCustomNamesEnumerableFieldClass>
    {
        public DefaultNoMatchingOptionalCustomNamesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoMatchingOptionalEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomNoMatchingOptionalEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("Year 2023", "NoSuchColumn"))
                .MakeOptional();
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingOptionalCustomNamesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesEnumerablePropertyClass>();
        Assert.Null(row1.CustomName);   
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingOptionalCustomNamesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingOptionalCustomNamesEnumerablePropertyClassMap>();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesEnumerablePropertyClass>();
        Assert.Null(row1.CustomName);   
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingOptionalEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoneMatchingOptionalEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoneMatchingOptionalCustomNamesEnumerablePropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "NoSuchColumn", "NoSuchColumn"}])]
        [ExcelOptional]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultNoneMatchingOptionalCustomNamesEnumerablePropertyClassMap : ExcelClassMap<NoneMatchingCustomNamesEnumerablePropertyClass>
    {
        public DefaultNoneMatchingOptionalCustomNamesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoneMatchingOptionalEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomNoneMatchingOptionalEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("NoSuchColumn", "NoSuchColumn"))
                .MakeOptional();
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingOptionalCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesEnumerableFieldClass>();
        Assert.Null(row1.CustomName);   
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingOptionalCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingOptionalCustomNamesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesEnumerableFieldClass>();
        Assert.Null(row1.CustomName);   
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingOptionalEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoneMatchingOptionalEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoneMatchingOptionalCustomNamesEnumerableFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "NoSuchColumn", "NoSuchColumn"}])]
        [ExcelOptional]
        public object?[] CustomName = default!;
    }

    private class DefaultNoneMatchingOptionalCustomNamesEnumerableFieldClassMap : ExcelClassMap<NoneMatchingOptionalCustomNamesEnumerableFieldClass>
    {
        public DefaultNoneMatchingOptionalCustomNamesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoneMatchingOptionalEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomNoneMatchingOptionalEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnsMatching(new NamesColumnMatcher("NoSuchColumn", "NoSuchColumn"))
                .MakeOptional();
        }
    }

    private class NamesColumnMatcher : IExcelColumnMatcher
    {
        public string[] ColumnNames { get; }

        public NamesColumnMatcher(params string[] columnNames)
        {
            ColumnNames = columnNames;
        }

        public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
            => ColumnNames.Contains(sheet.Heading!.GetColumnName(columnIndex));
    }
    
    
    [Fact]
    public void ReadRows_AutoMappedCustomNamesDictionaryProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesDictionaryPropertyClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedCustomNamesDictionaryProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNamesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesDictionaryPropertyClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }
    
    [Fact]
    public void ReadRows_CustomMappedDictionaryProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryPropertyClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }

    private class CustomNamesDictionaryPropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Column1", "Column2" }])]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DictionaryPropertyClass
    {
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultCustomNamesDictionaryPropertyClassMap : ExcelClassMap<CustomNamesDictionaryPropertyClass>
    {
        public DefaultCustomNamesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }
    
    private class CustomDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("Column1", "Column2"));
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomNamesDictionaryField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesDictionaryFieldClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedCustomNamesDictionaryField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNamesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesDictionaryFieldClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }
    
    [Fact]
    public void ReadRows_CustomMappedDictionaryField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryFieldClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }

    private class CustomNamesDictionaryFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Column1", "Column2" }])]
        public IDictionary<string, int> Value = default!;
    }

    private class DictionaryFieldClass
    {
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultCustomNamesDictionaryFieldClassMap : ExcelClassMap<CustomNamesDictionaryFieldClass>
    {
        public DefaultCustomNamesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }
    
    private class CustomDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("Column1", "Column2"));
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNamesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingCustomNamesDictionaryPropertyClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNamesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomNamesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingCustomNamesDictionaryPropertyClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingDictionaryProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryPropertyClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }

    private class NoMatchingCustomNamesDictionaryPropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Column1", "NoSuchColumn" }])]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultNoMatchingCustomNamesDictionaryPropertyClassMap : ExcelClassMap<NoMatchingCustomNamesDictionaryPropertyClass>
    {
        public DefaultNoMatchingCustomNamesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoMatchingDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomNoMatchingDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("Column1", "NoSuchColumn"));
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNamesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingCustomNamesDictionaryFieldClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNamesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomNamesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingCustomNamesDictionaryFieldClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryFieldClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }

    private class NoMatchingCustomNamesDictionaryFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Column1", "NoSuchColumn" }])]
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultNoMatchingCustomNamesDictionaryFieldClassMap : ExcelClassMap<NoMatchingCustomNamesDictionaryFieldClass>
    {
        public DefaultNoMatchingCustomNamesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoMatchingDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomNoMatchingDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("Column1", "NoSuchColumn"));
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomNamesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesDictionaryPropertyClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomNamesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingCustomNamesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesDictionaryPropertyClass>());
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoneMatchingDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoneMatchingDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryPropertyClass>());
    }

    private class NoneMatchingCustomNamesDictionaryPropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "NoSuchColumn", "NoSuchColumn" }])]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultNoneMatchingCustomNamesDictionaryPropertyClassMap : ExcelClassMap<NoneMatchingCustomNamesDictionaryPropertyClass>
    {
        public DefaultNoneMatchingCustomNamesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoneMatchingDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomNoneMatchingDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("NoSuchColumn", "NoSuchColumn"));
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomNamesDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesDictionaryFieldClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomNamesDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingCustomNamesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesDictionaryFieldClass>());
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoneMatchingDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoneMatchingDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryFieldClass>());
    }

    private class NoneMatchingCustomNamesDictionaryFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "NoSuchColumn", "NoSuchColumn" }])]
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultNoneMatchingCustomNamesDictionaryFieldClassMap : ExcelClassMap<NoneMatchingCustomNamesDictionaryFieldClass>
    {
        public DefaultNoneMatchingCustomNamesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoneMatchingDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomNoneMatchingDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("NoSuchColumn", "NoSuchColumn"));
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNamesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesDictionaryPropertyClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNamesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomNamesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesDictionaryPropertyClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingOptionalDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryPropertyClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }

    private class NoMatchingOptionalCustomNamesDictionaryPropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Column1", "NoSuchColumn" }])]
        [ExcelOptional]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultNoMatchingOptionalCustomNamesDictionaryPropertyClassMap : ExcelClassMap<NoMatchingOptionalCustomNamesDictionaryPropertyClass>
    {
        public DefaultNoMatchingOptionalCustomNamesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoMatchingOptionalDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomNoMatchingOptionalDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("Column1", "NoSuchColumn"))
                .MakeOptional();
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNamesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesDictionaryFieldClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNamesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomNamesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesDictionaryFieldClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingOptionalDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryFieldClass>();
        Assert.Single(row1.Value);
        Assert.Equal(1, row1.Value["Column1"]);
    }

    private class NoMatchingOptionalCustomNamesDictionaryFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "Column1", "NoSuchColumn" }])]
        [ExcelOptional]
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultNoMatchingOptionalCustomNamesDictionaryFieldClassMap : ExcelClassMap<NoMatchingOptionalCustomNamesDictionaryFieldClass>
    {
        public DefaultNoMatchingOptionalCustomNamesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoMatchingOptionalDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomNoMatchingOptionalDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("Column1", "NoSuchColumn"))
                .MakeOptional();
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoneMatchingOptionalCustomNamesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesDictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingOptionalCustomNamesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingOptionalCustomNamesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesDictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoneMatchingOptionalDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoneMatchingOptionalDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }

    private class NoneMatchingOptionalCustomNamesDictionaryPropertyClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "NoSuchColumn", "NoSuchColumn" }])]
        [ExcelOptional]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultNoneMatchingOptionalCustomNamesDictionaryPropertyClassMap : ExcelClassMap<NoneMatchingOptionalCustomNamesDictionaryPropertyClass>
    {
        public DefaultNoneMatchingOptionalCustomNamesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoneMatchingOptionalDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomNoneMatchingOptionalDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("NoSuchColumn", "NoSuchColumn"))
                .MakeOptional();
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoneMatchingOptionalCustomNamesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingOptionalCustomNamesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingOptionalCustomNamesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoneMatchingOptionalDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoneMatchingOptionalDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryFieldClass>();
        Assert.Null(row1.Value);
    }

    private class NoneMatchingOptionalCustomNamesDictionaryFieldClass
    {
        [ExcelColumnsMatching(typeof(NamesColumnMatcher), ConstructorArguments = [new string[] { "NoSuchColumn", "NoSuchColumn" }])]
        [ExcelOptional]
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultNoneMatchingOptionalCustomNamesDictionaryFieldClassMap : ExcelClassMap<NoneMatchingOptionalCustomNamesDictionaryFieldClass>
    {
        public DefaultNoneMatchingOptionalCustomNamesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoneMatchingOptionalDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomNoneMatchingOptionalDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnsMatching(new NamesColumnMatcher("NoSuchColumn", "NoSuchColumn"))
                .MakeOptional();
        }
    }
}
