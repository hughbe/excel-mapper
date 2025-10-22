namespace ExcelMapper.Tests;

public class MapCustomColumnIndicesTests
{
    [Fact]
    public void ReadRows_AutoMappedCustomIndicesMultipleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesMultiplePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndicesMultipleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesMultiplePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomIndicesMultiplePropertyClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        [ExcelColumnIndex(1)]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultCustomIndicesMultiplePropertyClassMap : ExcelClassMap<CustomIndicesMultiplePropertyClass>
    {
        public DefaultCustomIndicesMultiplePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndicesMultipleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesMultipleFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndicesMultipleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesMultipleFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomIndicesMultipleFieldClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        [ExcelColumnIndex(1)]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultCustomIndicesMultipleFieldClassMap : ExcelClassMap<CustomIndicesMultipleFieldClass>
    {
        public DefaultCustomIndicesMultipleFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndicesSingleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesSinglePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndicesSingleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndicesSinglePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesSinglePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomIndicesSinglePropertyClass
    {
        [ExcelColumnIndices(int.MaxValue, 1)]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultCustomIndicesSinglePropertyClassMap : ExcelClassMap<CustomIndicesSinglePropertyClass>
    {
        public DefaultCustomIndicesSinglePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndicesSingleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesSingleFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndicesSingleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndicesSingleFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesSingleFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomIndicesSingleFieldClass
    {
        [ExcelColumnIndices(int.MaxValue, 1)]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultCustomIndicesSingleFieldClassMap : ExcelClassMap<CustomIndicesSingleFieldClass>
    {
        public DefaultCustomIndicesSingleFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndicesEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesEnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndicesEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndicesEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesEnumerablePropertyClass>();
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

    private class CustomIndicesEnumerablePropertyClass
    {
        [ExcelColumnIndices(0, 1)]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultCustomIndicesEnumerablePropertyClassMap : ExcelClassMap<CustomIndicesEnumerablePropertyClass>
    {
        public DefaultCustomIndicesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class EnumerablePropertyClass
    {
        public object?[] CustomName { get; set; } = default!;
    }

    private class CustomEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(0, 1);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndicesEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesEnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndicesEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndicesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesEnumerableFieldClass>();
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

    private class CustomIndicesEnumerableFieldClass
    {
        [ExcelColumnIndices(0, 1)]
        public object?[] CustomName = default!;
    }

    private class DefaultCustomIndicesEnumerableFieldClassMap : ExcelClassMap<CustomIndicesEnumerableFieldClass>
    {
        public DefaultCustomIndicesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class EnumerableFieldClass
    {
        public object?[] CustomName = default!;
    }

    private class CustomEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(0, 1);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomIndicesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndicesEnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomIndicesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomIndicesEnumerablePropertyClassMap>();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndicesEnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<CustomNoMatchingEnumerablePropertyClassMap>();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerablePropertyClass>());
    }

    private class NoMatchingCustomIndicesEnumerablePropertyClass
    {
        [ExcelColumnIndices(0, int.MaxValue)]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultNoMatchingCustomIndicesEnumerablePropertyClassMap : ExcelClassMap<NoMatchingCustomIndicesEnumerablePropertyClass>
    {
        public DefaultNoMatchingCustomIndicesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }
    private class CustomNoMatchingEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomNoMatchingEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(0, int.MaxValue);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomIndicesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndicesEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomIndicesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomIndicesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndicesEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerableFieldClass>());
    }

    private class NoMatchingCustomIndicesEnumerableFieldClass
    {
        [ExcelColumnIndices(0, int.MaxValue)]
        public object?[] CustomName = default!;
    }

    private class DefaultNoMatchingCustomIndicesEnumerableFieldClassMap : ExcelClassMap<NoMatchingCustomIndicesEnumerableFieldClass>
    {
        public DefaultNoMatchingCustomIndicesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoMatchingEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomNoMatchingEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(0, int.MaxValue);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomIndicesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomIndicesEnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomIndicesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingCustomIndicesEnumerablePropertyClassMap>();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomIndicesEnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<CustomNoneMatchingEnumerablePropertyClassMap>();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerablePropertyClass>());
    }

    private class NoneMatchingCustomIndicesEnumerablePropertyClass
    {
        [ExcelColumnIndices(int.MaxValue, int.MaxValue)]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultNoneMatchingCustomIndicesEnumerablePropertyClassMap : ExcelClassMap<NoneMatchingCustomIndicesEnumerablePropertyClass>
    {
        public DefaultNoneMatchingCustomIndicesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }
    private class CustomNoneMatchingEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomNoneMatchingEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(int.MaxValue, int.MaxValue);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomIndicesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomIndicesEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomIndicesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingCustomIndicesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomIndicesEnumerableFieldClass>());
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

    private class NoneMatchingCustomIndicesEnumerableFieldClass
    {
        [ExcelColumnIndices(int.MaxValue, int.MaxValue)]
        public object?[] CustomName = default!;
    }

    private class DefaultNoneMatchingCustomIndicesEnumerableFieldClassMap : ExcelClassMap<NoneMatchingCustomIndicesEnumerableFieldClass>
    {
        public DefaultNoneMatchingCustomIndicesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoneMatchingEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomNoneMatchingEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(int.MaxValue, int.MaxValue);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomIndicesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndicesEnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomIndicesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomIndicesEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndicesEnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingOptionalEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoMatchingOptionalCustomIndicesEnumerablePropertyClass
    {
        [ExcelColumnIndices(0, int.MaxValue)]
        [ExcelOptional]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultNoMatchingOptionalCustomIndicesEnumerablePropertyClassMap : ExcelClassMap<NoMatchingOptionalCustomIndicesEnumerablePropertyClass>
    {
        public DefaultNoMatchingOptionalCustomIndicesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoMatchingOptionalEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomNoMatchingOptionalEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(0, int.MaxValue)
                .MakeOptional();
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomIndicesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndicesEnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomIndicesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomIndicesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndicesEnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingOptionalEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoMatchingOptionalCustomIndicesEnumerableFieldClass
    {
        [ExcelColumnIndices(0, int.MaxValue)]
        [ExcelOptional]
        public object?[] CustomName = default!;
    }

    private class DefaultNoMatchingOptionalCustomIndicesEnumerableFieldClassMap : ExcelClassMap<NoMatchingOptionalCustomIndicesEnumerableFieldClass>
    {
        public DefaultNoMatchingOptionalCustomIndicesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoMatchingOptionalEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomNoMatchingOptionalEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(0, int.MaxValue)
                .MakeOptional();
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingOptionalCustomIndicesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomIndicesEnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingOptionalCustomIndicesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingOptionalCustomIndicesEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomIndicesEnumerablePropertyClass>();
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

    private class NoneMatchingOptionalCustomIndicesEnumerablePropertyClass
    {
        [ExcelColumnIndices(int.MaxValue, int.MaxValue)]
        [ExcelOptional]
        public object?[] CustomName { get; set; } = default!;
    }

    private class DefaultNoneMatchingOptionalCustomIndicesEnumerablePropertyClassMap : ExcelClassMap<NoneMatchingOptionalCustomIndicesEnumerablePropertyClass>
    {
        public DefaultNoneMatchingOptionalCustomIndicesEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoneMatchingOptionalEnumerablePropertyClassMap : ExcelClassMap<EnumerablePropertyClass>
    {
        public CustomNoneMatchingOptionalEnumerablePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(int.MaxValue, int.MaxValue)
                .MakeOptional();
        }
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingOptionalCustomIndicesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomIndicesEnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingOptionalCustomIndicesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingOptionalCustomIndicesEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomIndicesEnumerableFieldClass>();
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

    private class NoneMatchingOptionalCustomIndicesEnumerableFieldClass
    {
        [ExcelColumnIndices(int.MaxValue, int.MaxValue)]
        [ExcelOptional]
        public object?[] CustomName = default!;
    }

    private class DefaultNoneMatchingOptionalCustomIndicesEnumerableFieldClassMap : ExcelClassMap<NoneMatchingOptionalCustomIndicesEnumerableFieldClass>
    {
        public DefaultNoneMatchingOptionalCustomIndicesEnumerableFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomNoneMatchingOptionalEnumerableFieldClassMap : ExcelClassMap<EnumerableFieldClass>
    {
        public CustomNoneMatchingOptionalEnumerableFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnIndices(int.MaxValue, int.MaxValue)
                .MakeOptional();
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomIndicesDictionaryProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesDictionaryPropertyClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedCustomIndicesDictionaryProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndicesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesDictionaryPropertyClass>();
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

    private class CustomIndicesDictionaryPropertyClass
    {
        [ExcelColumnIndices(1, 3)]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DictionaryPropertyClass
    {
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultCustomIndicesDictionaryPropertyClassMap : ExcelClassMap<CustomIndicesDictionaryPropertyClass>
    {
        public DefaultCustomIndicesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }
    
    private class CustomDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(1, 3);
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomIndicesDictionaryField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesDictionaryFieldClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedCustomIndicesDictionaryField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndicesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndicesDictionaryFieldClass>();
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

    private class CustomIndicesDictionaryFieldClass
    {
        [ExcelColumnIndices(1, 3)]
        public IDictionary<string, int> Value = default!;
    }

    private class DictionaryFieldClass
    {
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultCustomIndicesDictionaryFieldClassMap : ExcelClassMap<CustomIndicesDictionaryFieldClass>
    {
        public DefaultCustomIndicesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }
    
    private class CustomDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(1, 3);
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomIndicesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndicesDictionaryPropertyClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomIndicesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomIndicesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndicesDictionaryPropertyClass>());
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryPropertyClass>());
    }

    private class NoMatchingCustomIndicesDictionaryPropertyClass
    {
        [ExcelColumnIndices(1, int.MaxValue)]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultNoMatchingCustomIndicesDictionaryPropertyClassMap : ExcelClassMap<NoMatchingCustomIndicesDictionaryPropertyClass>
    {
        public DefaultNoMatchingCustomIndicesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoMatchingDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomNoMatchingDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(1, int.MaxValue);
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomIndicesDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndicesDictionaryFieldClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomIndicesDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomIndicesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndicesDictionaryFieldClass>());
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryFieldClass>());
    }

    private class NoMatchingCustomIndicesDictionaryFieldClass
    {
        [ExcelColumnIndices(1, int.MaxValue)]
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultNoMatchingCustomIndicesDictionaryFieldClassMap : ExcelClassMap<NoMatchingCustomIndicesDictionaryFieldClass>
    {
        public DefaultNoMatchingCustomIndicesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoMatchingDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomNoMatchingDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(1, int.MaxValue);
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomIndicesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomIndicesDictionaryPropertyClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomIndicesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingCustomIndicesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomIndicesDictionaryPropertyClass>());
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

    private class NoneMatchingCustomIndicesDictionaryPropertyClass
    {
        [ExcelColumnIndices(int.MaxValue, int.MaxValue)]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultNoneMatchingCustomIndicesDictionaryPropertyClassMap : ExcelClassMap<NoneMatchingCustomIndicesDictionaryPropertyClass>
    {
        public DefaultNoneMatchingCustomIndicesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoneMatchingDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomNoneMatchingDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(int.MaxValue, int.MaxValue);
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomIndicesDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomIndicesDictionaryFieldClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomIndicesDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingCustomIndicesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomIndicesDictionaryFieldClass>());
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

    private class NoneMatchingCustomIndicesDictionaryFieldClass
    {
        [ExcelColumnIndices(int.MaxValue, int.MaxValue)]
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultNoneMatchingCustomIndicesDictionaryFieldClassMap : ExcelClassMap<NoneMatchingCustomIndicesDictionaryFieldClass>
    {
        public DefaultNoneMatchingCustomIndicesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoneMatchingDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomNoneMatchingDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(int.MaxValue, int.MaxValue);
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomIndicesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndicesDictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomIndicesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomIndicesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndicesDictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingOptionalDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }

    private class NoMatchingOptionalCustomIndicesDictionaryPropertyClass
    {
        [ExcelColumnIndices(1, int.MaxValue)]
        [ExcelOptional]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultNoMatchingOptionalCustomIndicesDictionaryPropertyClassMap : ExcelClassMap<NoMatchingOptionalCustomIndicesDictionaryPropertyClass>
    {
        public DefaultNoMatchingOptionalCustomIndicesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoMatchingOptionalDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomNoMatchingOptionalDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(1, int.MaxValue)
                .MakeOptional();
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomIndicesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndicesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomIndicesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomIndicesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndicesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNoMatchingOptionalDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryFieldClass>();
        Assert.Null(row1.Value);
    }

    private class NoMatchingOptionalCustomIndicesDictionaryFieldClass
    {
        [ExcelColumnIndices(1, int.MaxValue)]
        [ExcelOptional]
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultNoMatchingOptionalCustomIndicesDictionaryFieldClassMap : ExcelClassMap<NoMatchingOptionalCustomIndicesDictionaryFieldClass>
    {
        public DefaultNoMatchingOptionalCustomIndicesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoMatchingOptionalDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomNoMatchingOptionalDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(1, int.MaxValue)
                .MakeOptional();
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoneMatchingOptionalCustomIndicesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomIndicesDictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingOptionalCustomIndicesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingOptionalCustomIndicesDictionaryPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomIndicesDictionaryPropertyClass>();
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

    private class NoneMatchingOptionalCustomIndicesDictionaryPropertyClass
    {
        [ExcelColumnIndices(int.MaxValue, int.MaxValue)]
        [ExcelOptional]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultNoneMatchingOptionalCustomIndicesDictionaryPropertyClassMap : ExcelClassMap<NoneMatchingOptionalCustomIndicesDictionaryPropertyClass>
    {
        public DefaultNoneMatchingOptionalCustomIndicesDictionaryPropertyClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoneMatchingOptionalDictionaryPropertyClassMap : ExcelClassMap<DictionaryPropertyClass>
    {
        public CustomNoneMatchingOptionalDictionaryPropertyClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(int.MaxValue, int.MaxValue)
                .MakeOptional();
        }
    }
    
    [Fact]
    public void ReadRows_AutoMappedNoneMatchingOptionalCustomIndicesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomIndicesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingOptionalCustomIndicesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoneMatchingOptionalCustomIndicesDictionaryFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomIndicesDictionaryFieldClass>();
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

    private class NoneMatchingOptionalCustomIndicesDictionaryFieldClass
    {
        [ExcelColumnIndices(int.MaxValue, int.MaxValue)]
        [ExcelOptional]
        public IDictionary<string, int> Value = default!;
    }

    private class DefaultNoneMatchingOptionalCustomIndicesDictionaryFieldClassMap : ExcelClassMap<NoneMatchingOptionalCustomIndicesDictionaryFieldClass>
    {
        public DefaultNoneMatchingOptionalCustomIndicesDictionaryFieldClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNoneMatchingOptionalDictionaryFieldClassMap : ExcelClassMap<DictionaryFieldClass>
    {
        public CustomNoneMatchingOptionalDictionaryFieldClassMap()
        {
            Map(o => o.Value)
                .WithColumnIndices(int.MaxValue, int.MaxValue)
                .MakeOptional();
        }
    }
}
