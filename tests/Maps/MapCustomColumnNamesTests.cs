namespace ExcelMapper.Tests;

public class MapCustomColumnNamesTests
{
    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultiplePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesMultiplePropertyClass
    {
        [ExcelColumnName("NoSuchColumn")]
        [ExcelColumnName("StringValue")]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultiplePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesMultipleFieldClass
    {
        [ExcelColumnName("NoSuchColumn")]
        [ExcelColumnName("StringValue")]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesMultipleOrdinalIgnoreCasePropertyClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.OrdinalIgnoreCase)]
        [ExcelColumnName("StringValue", StringComparison.OrdinalIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesMultipleOrdinalIgnoreCaseFieldClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.OrdinalIgnoreCase)]
        [ExcelColumnName("StringValue", StringComparison.OrdinalIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalPropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesMultipleOrdinalPropertyClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.Ordinal)]
        [ExcelColumnName("StringValue", StringComparison.Ordinal)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalPropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesMultipleOrdinalFieldClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.Ordinal)]
        [ExcelColumnName("StringValue", StringComparison.Ordinal)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleOrdinalFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleOrdinalFieldClass>());
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesMultipleCurrentCultureIgnoreCasePropertyClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.CurrentCultureIgnoreCase)]
        [ExcelColumnName("StringValue", StringComparison.CurrentCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesMultipleCurrentCultureIgnoreCaseFieldClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.CurrentCultureIgnoreCase)]
        [ExcelColumnName("StringValue", StringComparison.CurrentCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesMultipleCurrentCulturePropertyClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.CurrentCulture)]
        [ExcelColumnName("StringValue", StringComparison.CurrentCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesMultipleCurrentCultureFieldClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.CurrentCulture)]
        [ExcelColumnName("StringValue", StringComparison.CurrentCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleCurrentCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleCurrentCultureFieldClass>());
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesMultipleInvariantCultureIgnoreCasePropertyClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.InvariantCultureIgnoreCase)]
        [ExcelColumnName("StringValue", StringComparison.InvariantCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesMultipleInvariantCultureIgnoreCaseFieldClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.InvariantCultureIgnoreCase)]
        [ExcelColumnName("StringValue", StringComparison.InvariantCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesMultipleInvariantCulturePropertyClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.InvariantCulture)]
        [ExcelColumnName("StringValue", StringComparison.InvariantCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesMultipleInvariantCultureFieldClass
    {
        [ExcelColumnName("NoSuchColumn", StringComparison.InvariantCulture)]
        [ExcelColumnName("StringValue", StringComparison.InvariantCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesMultipleInvariantCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesMultipleInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesMultipleInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesMultipleInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSinglePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesSinglePropertyClass
    {
        [ExcelColumnNames("NoSuchColumn", "StringValue")]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSinglePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSinglePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesSingleFieldClass
    {
        [ExcelColumnNames("NoSuchColumn", "StringValue")]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesSingleOrdinalIgnoreCasePropertyClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.OrdinalIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleOrdinalIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesSingleOrdinalIgnoreCaseFieldClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.OrdinalIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleOrdinalIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleOrdinalIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleOrdinalIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalPropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesSingleOrdinalPropertyClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.Ordinal)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleOrdinalPropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalPropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesSingleOrdinalFieldClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.Ordinal)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleOrdinalFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleOrdinalFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleOrdinalPropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleOrdinalFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesSingleCurrentCultureIgnoreCasePropertyClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.CurrentCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleCurrentCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesSingleCurrentCultureIgnoreCaseFieldClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.CurrentCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleCurrentCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleCurrentCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleCurrentCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesSingleCurrentCulturePropertyClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.CurrentCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleCurrentCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesSingleCurrentCultureFieldClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.CurrentCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleCurrentCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleCurrentCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleCurrentCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleCurrentCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesSingleInvariantCultureIgnoreCasePropertyClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.InvariantCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleInvariantCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesSingleInvariantCultureIgnoreCaseFieldClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.InvariantCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleInvariantCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleInvariantCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleInvariantCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamesSingleInvariantCulturePropertyClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.InvariantCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleInvariantCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    private class CustomNamesSingleInvariantCultureFieldClass
    {
        [ExcelColumnNames(["NoSuchColumn", "StringValue"], StringComparison.InvariantCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleInvariantCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesSingleInvariantCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleInvariantCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesSingleInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesSingleInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesSingleInvariantCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesSingleInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerablePropertyClass
    {
        [ExcelColumnNames("Year 2023", "Year 2024")]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("Year 2023", "Year 2024");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
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

    private class CustomNamesEnumerableFieldClass
    {
        [ExcelColumnNames("Year 2023", "Year 2024")]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("Year 2023", "Year 2024");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class EnumerableFieldClass
    {
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableOrdinalIgnoreCasePropertyClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.OrdinalIgnoreCase)]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableOrdinalIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.OrdinalIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableOrdinalIgnoreCaseFieldClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.OrdinalIgnoreCase)]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableOrdinalIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.OrdinalIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableOrdinalIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.OrdinalIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableOrdinalIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.OrdinalIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalPropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableOrdinalPropertyClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.Ordinal)]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableOrdinalPropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalPropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.Ordinal);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableOrdinalFieldClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.Ordinal)]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableOrdinalFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableOrdinalFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.Ordinal);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableOrdinalPropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.Ordinal);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableOrdinalFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.Ordinal);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableCurrentCultureIgnoreCasePropertyClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCultureIgnoreCase)]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableCurrentCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCultureIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableCurrentCultureIgnoreCaseFieldClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCultureIgnoreCase)]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableCurrentCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCultureIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableCurrentCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCultureIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableCurrentCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCultureIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableCurrentCulturePropertyClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCulture)]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableCurrentCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCulture);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableCurrentCultureFieldClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCulture)]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableCurrentCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCulture);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableCurrentCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCulture);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableCurrentCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.CurrentCulture);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableInvariantCultureIgnoreCasePropertyClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCultureIgnoreCase)]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableInvariantCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCultureIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableInvariantCultureIgnoreCaseFieldClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCultureIgnoreCase)]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableInvariantCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCultureIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableInvariantCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCultureIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableInvariantCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCultureIgnoreCase);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableInvariantCulturePropertyClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCulture)]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableInvariantCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCulture);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    private class CustomNamesEnumerableInvariantCultureFieldClass
    {
        [ExcelColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCulture)]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableInvariantCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamesEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCulture);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Equal(["1", "2"], row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableInvariantCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCulture);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamesEnumerableInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamesEnumerableInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamesEnumerableInvariantCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNamesEnumerableInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedEnumerableInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames(["Year 2023", "Year 2024"], StringComparison.InvariantCulture);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNamesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamesEnumerablePropertyClass>());
    }

    private class NoMatchingCustomNamesEnumerablePropertyClass
    {
        [ExcelColumnNames("Year 2023", "NoSuchColumn")]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNamesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<NoMatchingCustomNamesEnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamesEnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("Year 2023", "NoSuchColumn");
        });
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNamesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamesEnumerableFieldClass>());
    }

    private class NoMatchingCustomNamesEnumerableFieldClass
    {
        [ExcelColumnNames("Year 2023", "NoSuchColumn")]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNamesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingCustomNamesEnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamesEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("Year 2023", "NoSuchColumn");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomNamesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesEnumerablePropertyClass>());
    }

    private class NoneMatchingCustomNamesEnumerablePropertyClass
    {
        [ExcelColumnNames("NoSuchColumn", "NoSuchColumn")]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomNamesEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<NoneMatchingCustomNamesEnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesEnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingEnumerableProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("NoSuchColumn", "NoSuchColumn");
        });
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerablePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoneMatchingCustomNamesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesEnumerableFieldClass>());
    }

    private class NoneMatchingCustomNamesEnumerableFieldClass
    {
        [ExcelColumnNames("NoSuchColumn", "NoSuchColumn")]
        public object?[] CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedNoneMatchingCustomNamesEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<NoneMatchingCustomNamesEnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingEnumerableField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("NoSuchColumn", "NoSuchColumn");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNamesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesEnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNamesEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingOptionalCustomNamesEnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesEnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("Year 2023", "NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoMatchingOptionalCustomNamesEnumerablePropertyClass
    {
        [ExcelColumnNames("Year 2023", "NoSuchColumn")]
        [ExcelOptional]
        public object?[] CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesEnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNamesEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingOptionalCustomNamesEnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesEnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("Year 2023", "NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoMatchingOptionalCustomNamesEnumerableFieldClass
    {
        [ExcelColumnNames("Year 2023", "NoSuchColumn")]
        [ExcelOptional]
        public object?[] CustomName = default!;
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
        importer.Configuration.RegisterClassMap<NoneMatchingOptionalCustomNamesEnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesEnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingOptionalEnumerableProperty_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("NoSuchColumn", "NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerablePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoneMatchingOptionalCustomNamesEnumerablePropertyClass
    {
        [ExcelColumnNames("NoSuchColumn", "NoSuchColumn")]
        [ExcelOptional]
        public object?[] CustomName { get; set; } = default!;
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
        importer.Configuration.RegisterClassMap<NoneMatchingOptionalCustomNamesEnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesEnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedNoneMatchingOptionalEnumerableField_Success()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<EnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnNames("NoSuchColumn", "NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EnumerableFieldClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoneMatchingOptionalCustomNamesEnumerableFieldClass
    {
        [ExcelColumnNames("NoSuchColumn", "NoSuchColumn")]
        [ExcelOptional]
        public object?[] CustomName = default!;
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
        importer.Configuration.RegisterClassMap<CustomNamesDictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value);
        });

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
        importer.Configuration.RegisterClassMap<DictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("Column1", "Column2");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryPropertyClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }

    private class CustomNamesDictionaryPropertyClass
    {
        [ExcelColumnNames("Column1", "Column2")]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DictionaryPropertyClass
    {
        public IDictionary<string, int> Value { get; set; } = default!;
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
        importer.Configuration.RegisterClassMap<CustomNamesDictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value);
        });

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
        importer.Configuration.RegisterClassMap<DictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("Column1", "Column2");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryFieldClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
    }

    private class CustomNamesDictionaryFieldClass
    {
        [ExcelColumnNames("Column1", "Column2")]
        public IDictionary<string, int> Value = default!;
    }

    private class DictionaryFieldClass
    {
        public IDictionary<string, int> Value = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNamesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamesDictionaryPropertyClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNamesDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingCustomNamesDictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamesDictionaryPropertyClass>());
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("Column1", "NoSuchColumn");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryPropertyClass>());
    }

    private class NoMatchingCustomNamesDictionaryPropertyClass
    {
        [ExcelColumnNames("Column1", "NoSuchColumn")]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNamesDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamesDictionaryFieldClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNamesDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingCustomNamesDictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamesDictionaryFieldClass>());
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("Column1", "NoSuchColumn");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryFieldClass>());
    }

    private class NoMatchingCustomNamesDictionaryFieldClass
    {
        [ExcelColumnNames("Column1", "NoSuchColumn")]
        public IDictionary<string, int> Value = default!;
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
        importer.Configuration.RegisterClassMap<NoneMatchingCustomNamesDictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesDictionaryPropertyClass>());
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoneMatchingDictionaryProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("NoSuchColumn", "NoSuchColumn");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryPropertyClass>());
    }

    private class NoneMatchingCustomNamesDictionaryPropertyClass
    {
        [ExcelColumnNames("NoSuchColumn", "NoSuchColumn")]
        public IDictionary<string, int> Value { get; set; } = default!;
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
        importer.Configuration.RegisterClassMap<NoneMatchingCustomNamesDictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoneMatchingCustomNamesDictionaryFieldClass>());
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoneMatchingDictionaryField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("NoSuchColumn", "NoSuchColumn");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryFieldClass>());
    }

    private class NoneMatchingCustomNamesDictionaryFieldClass
    {
        [ExcelColumnNames("NoSuchColumn", "NoSuchColumn")]
        public IDictionary<string, int> Value = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNamesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesDictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNamesDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingOptionalCustomNamesDictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesDictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("Column1", "NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }

    private class NoMatchingOptionalCustomNamesDictionaryPropertyClass
    {
        [ExcelColumnNames("Column1", "NoSuchColumn")]
        [ExcelOptional]
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNamesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNamesDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingOptionalCustomNamesDictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoMatchingOptionalDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("Column1", "NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryFieldClass>();
        Assert.Null(row1.Value);
    }

    private class NoMatchingOptionalCustomNamesDictionaryFieldClass
    {
        [ExcelColumnNames("Column1", "NoSuchColumn")]
        [ExcelOptional]
        public IDictionary<string, int> Value = default!;
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
        importer.Configuration.RegisterClassMap<NoneMatchingOptionalCustomNamesDictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesDictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoneMatchingOptionalDictionaryProperty_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryPropertyClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("NoSuchColumn", "NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryPropertyClass>();
        Assert.Null(row1.Value);
    }

    private class NoneMatchingOptionalCustomNamesDictionaryPropertyClass
    {
        [ExcelColumnNames("NoSuchColumn", "NoSuchColumn")]
        [ExcelOptional]
        public IDictionary<string, int> Value { get; set; } = default!;
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
        importer.Configuration.RegisterClassMap<NoneMatchingOptionalCustomNamesDictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }
    
    [Fact]
    public void ReadRows_CustomMappedNoneMatchingOptionalDictionaryField_Success()
    {
        using var importer = Helpers.GetImporter("CustomDictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<NoneMatchingOptionalCustomNamesDictionaryFieldClass>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnNames("NoSuchColumn", "NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoneMatchingOptionalCustomNamesDictionaryFieldClass>();
        Assert.Null(row1.Value);
    }

    private class NoneMatchingOptionalCustomNamesDictionaryFieldClass
    {
        [ExcelColumnNames("NoSuchColumn", "NoSuchColumn")]
        [ExcelOptional]
        public IDictionary<string, int> Value = default!;
    }
}
