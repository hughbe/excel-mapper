using Xunit;

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
    
    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomNameProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamePropertyClass>());
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
    public void ReadRows_AutoMappedCustomNameField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameFieldClass>();
        Assert.Equal("a", row1.CustomName);
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
    public void ReadRows_AutoMappedNoMatchingOptionalCustomNameField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNameFieldClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNamePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamePropertyClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNameProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomNamePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNamePropertyClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNameProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomNamePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNamePropertyClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNameFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameFieldClass>();
        Assert.Equal("a", row1.CustomName);
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomNameField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomNameFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomNameFieldClass>());
    }
    
    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomNameField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomNameFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomNameFieldClass>();
        Assert.Null(row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomNamePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNamePropertyClass>();
        Assert.Equal("1", row1.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomNameFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameFieldClass>();
        Assert.Equal("1", row1.CustomName);
    }

    private class CustomNamePropertyClass
    {
        [ExcelColumnName("StringValue")]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultCustomNamePropertyClassMap : ExcelClassMap<CustomNamePropertyClass>
    {
        public DefaultCustomNamePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class NoMatchingCustomNamePropertyClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultNoMatchingCustomNamePropertyClassMap : ExcelClassMap<NoMatchingCustomNamePropertyClass>
    {
        public DefaultNoMatchingCustomNamePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class NoMatchingOptionalCustomNamePropertyClass
    {
        [ExcelColumnName("NoSuchColumn")]
        [ExcelOptional]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultNoMatchingOptionalCustomNamePropertyClassMap : ExcelClassMap<NoMatchingOptionalCustomNamePropertyClass>
    {
        public DefaultNoMatchingOptionalCustomNamePropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomCustomNamePropertyClassMap : ExcelClassMap<CustomNamePropertyClass>
    {
        public CustomCustomNamePropertyClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnName("Int Value");
        }
    }
    
    private class CustomNameFieldClass
    {
        [ExcelColumnName("StringValue")]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultCustomNameFieldClassMap : ExcelClassMap<CustomNameFieldClass>
    {
        public DefaultCustomNameFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }
    
    private class NoMatchingCustomNameFieldClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultNoMatchingCustomNameFieldClassMap : ExcelClassMap<NoMatchingCustomNameFieldClass>
    {
        public DefaultNoMatchingCustomNameFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }
    
    private class NoMatchingOptionalCustomNameFieldClass
    {
        [ExcelColumnName("NoSuchColumn")]
        [ExcelOptional]
        public string CustomName { get; set; } = default!;
    }

    private class DefaultNoMatchingOptionalCustomNameFieldClassMap : ExcelClassMap<NoMatchingOptionalCustomNameFieldClass>
    {
        public DefaultNoMatchingOptionalCustomNameFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomCustomNameFieldClassMap : ExcelClassMap<CustomNameFieldClass>
    {
        public CustomCustomNameFieldClassMap()
        {
            Map(p => p.CustomName)
                .WithColumnName("Int Value");
        }
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

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNameEnumPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNameEnumFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomNameEnumPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        var row2 = sheet.ReadRow<CustomNameEnumPropertyClass>();
        Assert.Equal(CustomEnum.B, row2.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomNameEnumFieldClassMap>();

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

    [Fact]
    public void ReadRows_DefaultMappedCustomNameNullableEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNameNullableEnumPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomNameNullableEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNameNullableEnumFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameNullableEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomNameNullableEnumPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        var row2 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.B, row2.CustomName);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomNameNullableEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomNameNullableEnumFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomName);

        var row2 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.B, row2.CustomName);
    }

    private class CustomNameEnumPropertyClass
    {
        [ExcelColumnName("StringValue")]
        public CustomEnum CustomName { get; set; }
    }

    private class DefaultCustomNameEnumPropertyClassMap : ExcelClassMap<CustomNameEnumPropertyClass>
    {
        public DefaultCustomNameEnumPropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomCustomNameEnumPropertyClassMap : ExcelClassMap<CustomNameEnumPropertyClass>
    {
        public CustomCustomNameEnumPropertyClassMap()
        {
            Map(p => p.CustomName, ignoreCase: true);
        }
    }

    private class CustomNameNullableEnumPropertyClass
    {
        [ExcelColumnName("StringValue")]
        public CustomEnum? CustomName { get; set; }
    }

    private class DefaultCustomNameNullableEnumPropertyClassMap : ExcelClassMap<CustomNameNullableEnumPropertyClass>
    {
        public DefaultCustomNameNullableEnumPropertyClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomCustomNameNullableEnumPropertyClassMap : ExcelClassMap<CustomNameNullableEnumPropertyClass>
    {
        public CustomCustomNameNullableEnumPropertyClassMap()
        {
            Map(p => p.CustomName, ignoreCase: true);
        }
    }
    
    private class CustomNameEnumFieldClass
    {
        [ExcelColumnName("StringValue")]
        public CustomEnum CustomName { get; set; }
    }

    private class DefaultCustomNameEnumFieldClassMap : ExcelClassMap<CustomNameEnumFieldClass>
    {
        public DefaultCustomNameEnumFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomCustomNameEnumFieldClassMap : ExcelClassMap<CustomNameEnumFieldClass>
    {
        public CustomCustomNameEnumFieldClassMap()
        {
            Map(p => p.CustomName, ignoreCase: true);
        }
    }
    
    private class CustomNameNullableEnumFieldClass
    {
        [ExcelColumnName("StringValue")]
        public CustomEnum? CustomName { get; set; }
    }

    private class DefaultCustomNameNullableEnumFieldClassMap : ExcelClassMap<CustomNameNullableEnumFieldClass>
    {
        public DefaultCustomNameNullableEnumFieldClassMap()
        {
            Map(p => p.CustomName);
        }
    }

    private class CustomCustomNameNullableEnumFieldClassMap : ExcelClassMap<CustomNameNullableEnumFieldClass>
    {
        public CustomCustomNameNullableEnumFieldClassMap()
        {
            Map(p => p.CustomName, ignoreCase: true);
        }
    }

    private enum CustomEnum
    {
        a,
        B
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

    [Fact]
    public void ReadRows_DefaultMappedCustomNameEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNameEnumerablePropertyClassMap>();

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
    public void ReadRows_DefaultMappedCustomNameEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomNameEnumerableFieldClassMap>();

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

    private class CustomNameEnumerablePropertyClass
    {
        [ExcelColumnName("Value")]
        public object?[] CustomValue { get; set; } = default!;
    }

    private class DefaultCustomNameEnumerablePropertyClassMap : ExcelClassMap<CustomNameEnumerablePropertyClass>
    {
        public DefaultCustomNameEnumerablePropertyClassMap()
        {
            Map(p => p.CustomValue);
        }
    }
    
    private class CustomNameEnumerableFieldClass
    {
        [ExcelColumnName("Value")]
        public object?[] CustomValue = default!;
    }

    private class DefaultCustomNameEnumerableFieldClassMap : ExcelClassMap<CustomNameEnumerableFieldClass>
    {
        public DefaultCustomNameEnumerableFieldClassMap()
        {
            Map(p => p.CustomValue);
        }
    }
}
