namespace ExcelMapper.Tests;

public class MapCustomColumnIndexTests
{
    [Fact]
    public void ReadRows_AutoMappedCustomNameProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNamePropertyClass
    {
        [ExcelColumnName("StringValue")]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNamePropertyClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnName("Int Value");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamePropertyClass>();
        Assert.Equal("1", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameFieldClass
    {
        [ExcelColumnName("StringValue")]
        public string CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameFieldClass>(c =>
        {
            c.Map(p => p.CustomName)
                .WithColumnName("Int Value");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameFieldClass>();
        Assert.Equal("1", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyOrdinalIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameOrdinalIgnoreCasePropertyClass
    {
        [ExcelColumnName("StringValue", StringComparison.OrdinalIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyOrdinalIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameOrdinalIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldOrdinalIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameOrdinalIgnoreCaseFieldClass
    {
        [ExcelColumnName("StringValue", StringComparison.OrdinalIgnoreCase)]
        public string CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldOrdinalIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameOrdinalIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldOrdinalIgnoreCaseMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldOrdinalIgnoreCaseMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameOrdinalIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyOrdinal_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalPropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameOrdinalPropertyClass
    {
        [ExcelColumnName("StringValue", StringComparison.Ordinal)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyOrdinal_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameOrdinalPropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalPropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldOrdinal_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameOrdinalFieldClass
    {
        [ExcelColumnName("StringValue", StringComparison.Ordinal)]
        public string CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldOrdinal_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameOrdinalFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameOrdinalFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyOrdinalNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyOrdinalNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameOrdinalPropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldOrdinalNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldOrdinalNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameOrdinalFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyCurrentCultureIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameCurrentCultureIgnoreCasePropertyClass
    {
        [ExcelColumnName("StringValue", StringComparison.CurrentCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyCurrentCultureIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameCurrentCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldCurrentCultureIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameCurrentCultureIgnoreCaseFieldClass
    {
        [ExcelColumnName("StringValue", StringComparison.CurrentCultureIgnoreCase)]
        public string CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldCurrentCultureIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameCurrentCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldCurrentCultureIgnoreCaseMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldCurrentCultureIgnoreCaseMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameCurrentCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyCurrentCulture_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameCurrentCulturePropertyClass
    {
        [ExcelColumnName("StringValue", StringComparison.CurrentCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyCurrentCulture_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameCurrentCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldCurrentCulture_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameCurrentCultureFieldClass
    {
        [ExcelColumnName("StringValue", StringComparison.CurrentCulture)]
        public string CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldCurrentCulture_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameCurrentCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameCurrentCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyCurrentCultureNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyCurrentCultureNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameCurrentCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldCurrentCultureNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldCurrentCultureNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameCurrentCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyInvariantCultureIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameInvariantCultureIgnoreCasePropertyClass
    {
        [ExcelColumnName("StringValue", StringComparison.InvariantCultureIgnoreCase)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyInvariantCultureIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameInvariantCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldInvariantCultureIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameInvariantCultureIgnoreCaseFieldClass
    {
        [ExcelColumnName("StringValue", StringComparison.InvariantCultureIgnoreCase)]
        public string CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldInvariantCultureIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameInvariantCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldInvariantCultureIgnoreCaseMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldInvariantCultureIgnoreCaseMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameInvariantCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyInvariantCulture_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameInvariantCulturePropertyClass
    {
        [ExcelColumnName("StringValue", StringComparison.InvariantCulture)]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyInvariantCulture_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameInvariantCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCulturePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldInvariantCulture_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    private class CustomNameInvariantCultureFieldClass
    {
        [ExcelColumnName("StringValue", StringComparison.InvariantCulture)]
        public string CustomName = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldInvariantCulture_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameInvariantCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameInvariantCultureFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNamePropertyInvariantCultureNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNamePropertyInvariantCultureNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameInvariantCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameFieldInvariantCultureNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameFieldInvariantCultureNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnName_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameInvariantCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNameProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamePropertyClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNameProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingCustomNamePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNameField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNameFieldClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNameField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingCustomNameFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNameFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNameProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamePropertyClass>();
        Assert.Null(row1.CustomName);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNameProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingOptionalCustomNamePropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNameField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNameFieldClass>();
        Assert.Null(row1.CustomName);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNameField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoMatchingOptionalCustomNameFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNameFieldClass>();
        Assert.Null(row1.CustomName);
    }

    private class NoMatchingCustomNamePropertyClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public string CustomName { get; set; } = default!;
    }

    private class NoMatchingOptionalCustomNamePropertyClass
    {
        [ExcelColumnName("NoSuchColumn")]
        [ExcelOptional]
        public string CustomName { get; set; } = default!;
    }

    private class NoMatchingCustomNameFieldClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public string CustomName { get; set; } = default!;
    }
    
    private class NoMatchingOptionalCustomNameFieldClass
    {
        [ExcelColumnName("NoSuchColumn")]
        [ExcelOptional]
        public string CustomName { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumPropertyClass>());
    }

    private class CustomNameEnumPropertyClass
    {
        [ExcelColumnName("StringValue")]
        public CustomEnum CustomName { get; set; }
    }

    private enum CustomEnum
    {
        a,
        B
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumPropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumPropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumPropertyClass>(c =>
        {
            c.Map(p => p.CustomName, ignoreCase: true);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        var row2 = sheet.ReadRow<CustomNameEnumPropertyClass>();
        Assert.Equal(CustomEnum.B, row2.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumFieldClass>());
    }
    
    private class CustomNameEnumFieldClass
    {
        [ExcelColumnName("StringValue")]
        public CustomEnum CustomName { get; set; }
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumFieldClass>(c =>
        {
            c.Map(p => p.CustomName, ignoreCase: true);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        var row2 = sheet.ReadRow<CustomNameEnumFieldClass>();
        Assert.Equal(CustomEnum.B, row2.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameNullableEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumPropertyClass>());
    }

    private class CustomNameNullableEnumPropertyClass
    {
        [ExcelColumnName("StringValue")]
        public CustomEnum? CustomName { get; set; }
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameNullableEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameNullableEnumPropertyClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumPropertyClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameNullableEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameNullableEnumPropertyClass>(c =>
        {
            c.Map(p => p.CustomName, ignoreCase: true);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        var row2 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.B, row2.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameNullableEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumFieldClass>());
    }

    private class CustomNameNullableEnumFieldClass
    {
        [ExcelColumnName("StringValue")]
        public CustomEnum? CustomName { get; set; }
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameNullableEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameNullableEnumFieldClass>(c =>
        {
            c.Map(p => p.CustomName);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameNullableEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameNullableEnumFieldClass>(c =>
        {
            c.Map(p => p.CustomName, ignoreCase: true);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        var row2 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.B, row2.CustomName);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerablePropertyClass
    {
        [ExcelColumnName("Value")]
        public object?[] CustomValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerablePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }
    private class CustomNameEnumerableFieldClass
    {
        [ExcelColumnName("Value")]
        public object?[] CustomValue = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableOrdinalIgnoreCasePropertyClass
    {
        [ExcelColumnName("Value", StringComparison.OrdinalIgnoreCase)]
        public object?[] CustomValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableOrdinalIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableOrdinalIgnoreCaseFieldClass
    {
        [ExcelColumnName("Value", StringComparison.OrdinalIgnoreCase)]
        public object?[] CustomValue = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableOrdinalIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableOrdinalIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableOrdinalIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableOrdinalPropertyClass
    {
        [ExcelColumnName("Value", StringComparison.Ordinal)]
        public object?[] CustomValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableOrdinalProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableOrdinalPropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableOrdinalFieldClass
    {
        [ExcelColumnName("Value", StringComparison.Ordinal)]
        public object?[] CustomValue = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableOrdinalField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableOrdinalFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableOrdinalPropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableOrdinalPropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableOrdinalPropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableOrdinalFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableOrdinalFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableOrdinalFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass
    {
        [ExcelColumnName("Value", StringComparison.CurrentCultureIgnoreCase)]
        public object?[] CustomValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableCurrentCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass
    {
        [ExcelColumnName("Value", StringComparison.CurrentCultureIgnoreCase)]
        public object?[] CustomValue = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableCurrentCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableCurrentCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableCurrentCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableCurrentCulturePropertyClass
    {
        [ExcelColumnName("Value", StringComparison.CurrentCulture)]
        public object?[] CustomValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableCurrentCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableCurrentCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableCurrentCultureFieldClass
    {
        [ExcelColumnName("Value", StringComparison.CurrentCulture)]
        public object?[] CustomValue = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableCurrentCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableCurrentCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableCurrentCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableCurrentCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableCurrentCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableCurrentCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableCurrentCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableCurrentCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass
    {
        [ExcelColumnName("Value", StringComparison.InvariantCultureIgnoreCase)]
        public object?[] CustomValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableInvariantCultureIgnoreCaseProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass
    {
        [ExcelColumnName("Value", StringComparison.InvariantCultureIgnoreCase)]
        public object?[] CustomValue = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableInvariantCultureIgnoreCaseField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableInvariantCultureIgnoreCasePropertyMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCasePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableInvariantCultureIgnoreCaseFieldMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureIgnoreCaseFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableInvariantCulturePropertyClass
    {
        [ExcelColumnName("Value", StringComparison.InvariantCulture)]
        public object?[] CustomValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableInvariantCultureProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableInvariantCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomNameEnumerableInvariantCultureFieldClass
    {
        [ExcelColumnName("Value", StringComparison.InvariantCulture)]
        public object?[] CustomValue = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableInvariantCultureField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableInvariantCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableInvariantCulturePropertyNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableInvariantCulturePropertyClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableInvariantCulturePropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomNameEnumerableInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableInvariantCultureFieldNoMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_IgnoreCase.xlsx");
        importer.Configuration.RegisterClassMap<CustomNameEnumerableInvariantCultureFieldClass>(c =>
        {
            c.Map(p => p.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumerableInvariantCultureFieldClass>());
    }
}
