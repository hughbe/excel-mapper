using Xunit;

namespace ExcelMapper.Tests;

public class MapCustomColumnNameTests
{
    [Fact]
    public void ReadRows_AutoMappedCustomIndexProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexPropertyClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomIndexProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndexPropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomIndexProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndexPropertyClass>();
        Assert.Null(row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexFieldClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingCustomIndexField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndexFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingOptionalCustomIndexField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndexFieldClass>();
        Assert.Null(row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexSingleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexSinglePropertyClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexSingleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexSingleFieldClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexMultipleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexMultiplePropertyClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexMultipleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexMultipleFieldClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexPropertyClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomIndexProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomIndexPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndexPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomIndexProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomIndexPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndexPropertyClass>();
        Assert.Null(row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexFieldClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingCustomIndexField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingCustomIndexFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoMatchingCustomIndexFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedNoMatchingOptionalCustomIndexField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNoMatchingOptionalCustomIndexFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NoMatchingOptionalCustomIndexFieldClass>();
        Assert.Null(row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexSingleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexSinglePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexSinglePropertyClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexSingleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexSingleFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexSingleFieldClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexMultipleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexMultiplePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexMultiplePropertyClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexMultipleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexMultipleFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexMultipleFieldClass>();
        Assert.Equal("a", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomIndexProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomIndexPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexPropertyClass>();
        Assert.Equal("1", row1.CustomIndex);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomIndexField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomIndexFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexFieldClass>();
        Assert.Equal("1", row1.CustomIndex);
    }

    private class CustomIndexPropertyClass
    {
        [ExcelColumnIndex(1)]
        public string CustomIndex { get; set; } = default!;
    }

    private class DefaultCustomIndexPropertyClassMap : ExcelClassMap<CustomIndexPropertyClass>
    {
        public DefaultCustomIndexPropertyClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class CustomCustomIndexPropertyClassMap : ExcelClassMap<CustomIndexPropertyClass>
    {
        public CustomCustomIndexPropertyClassMap()
        {
            Map(p => p.CustomIndex)
                .WithColumnName("Int Value");
        }
    }

    private class NoMatchingCustomIndexPropertyClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        public string CustomIndex { get; set; } = default!;
    }

    private class DefaultNoMatchingCustomIndexPropertyClassMap : ExcelClassMap<NoMatchingCustomIndexPropertyClass>
    {
        public DefaultNoMatchingCustomIndexPropertyClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class NoMatchingOptionalCustomIndexPropertyClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        [ExcelOptional]
        public string CustomIndex { get; set; } = default!;
    }

    private class DefaultNoMatchingOptionalCustomIndexPropertyClassMap : ExcelClassMap<NoMatchingOptionalCustomIndexPropertyClass>
    {
        public DefaultNoMatchingOptionalCustomIndexPropertyClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class NoMatchingCustomIndexFieldClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        public string CustomIndex = default!;
    }

    private class DefaultNoMatchingCustomIndexFieldClassMap : ExcelClassMap<NoMatchingCustomIndexFieldClass>
    {
        public DefaultNoMatchingCustomIndexFieldClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class NoMatchingOptionalCustomIndexFieldClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        [ExcelOptional]
        public string CustomIndex = default!;
    }

    private class DefaultNoMatchingOptionalCustomIndexFieldClassMap : ExcelClassMap<NoMatchingOptionalCustomIndexFieldClass>
    {
        public DefaultNoMatchingOptionalCustomIndexFieldClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }
    
    private class CustomIndexFieldClass
    {
        [ExcelColumnIndex(1)]
        public string CustomIndex = default!;
    }

    private class DefaultCustomIndexFieldClassMap : ExcelClassMap<CustomIndexFieldClass>
    {
        public DefaultCustomIndexFieldClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class CustomCustomIndexFieldClassMap : ExcelClassMap<CustomIndexFieldClass>
    {
        public CustomCustomIndexFieldClassMap()
        {
            Map(p => p.CustomIndex)
                .WithColumnName("Int Value");
        }
    }

    private class CustomIndexMultiplePropertyClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        [ExcelColumnIndex(1)]
        public string CustomIndex { get; set; } = default!;
    }

    private class DefaultCustomIndexMultiplePropertyClassMap : ExcelClassMap<CustomIndexMultiplePropertyClass>
    {
        public DefaultCustomIndexMultiplePropertyClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }
    
    private class CustomIndexMultipleFieldClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        [ExcelColumnIndex(1)]
        public string CustomIndex = default!;
    }

    private class DefaultCustomIndexMultipleFieldClassMap : ExcelClassMap<CustomIndexMultipleFieldClass>
    {
        public DefaultCustomIndexMultipleFieldClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class CustomIndexSinglePropertyClass
    {
        [ExcelColumnIndices(int.MaxValue, 1)]
        public string CustomIndex { get; set; } = default!;
    }

    private class DefaultCustomIndexSinglePropertyClassMap : ExcelClassMap<CustomIndexSinglePropertyClass>
    {
        public DefaultCustomIndexSinglePropertyClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }
    
    private class CustomIndexSingleFieldClass
    {
        [ExcelColumnIndices(int.MaxValue, 1)]
        public string CustomIndex = default!;
    }

    private class DefaultCustomIndexSingleFieldClassMap : ExcelClassMap<CustomIndexSingleFieldClass>
    {
        public DefaultCustomIndexSingleFieldClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexEnumPropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexEnumFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexEnumPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexEnumFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedCustomIndexEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomIndexEnumPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        var row2 = sheet.ReadRow<CustomIndexEnumPropertyClass>();
        Assert.Equal(CustomEnum.B, row2.CustomIndex);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomIndexEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomIndexEnumFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        var row2 = sheet.ReadRow<CustomIndexEnumFieldClass>();
        Assert.Equal(CustomEnum.B, row2.CustomIndex);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexNullableEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexNullableEnumPropertyClass>());
    }


    [Fact]
    public void ReadRows_AutoMappedCustomIndexNullableEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexNullableEnumFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexNullableEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexNullableEnumPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexNullableEnumPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexNullableEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexNullableEnumFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexNullableEnumFieldClass>());
    }

    [Fact]
    public void ReadRows_CustomMappedCustomIndexNullableEnumProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomIndexNullableEnumPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        var row2 = sheet.ReadRow<CustomIndexNullableEnumPropertyClass>();
        Assert.Equal(CustomEnum.B, row2.CustomIndex);
    }

    [Fact]
    public void ReadRows_CustomMappedCustomIndexNullableEnumField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomIndexNullableEnumFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.a, row1.CustomIndex);

        var row2 = sheet.ReadRow<CustomIndexNullableEnumFieldClass>();
        Assert.Equal(CustomEnum.B, row2.CustomIndex);
    }

    private class CustomIndexEnumPropertyClass
    {
        [ExcelColumnIndex(1)]
        public CustomEnum CustomIndex { get; set; }
    }

    private class DefaultCustomIndexEnumPropertyClassMap : ExcelClassMap<CustomIndexEnumPropertyClass>
    {
        public DefaultCustomIndexEnumPropertyClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class CustomCustomIndexEnumPropertyClassMap : ExcelClassMap<CustomIndexEnumPropertyClass>
    {
        public CustomCustomIndexEnumPropertyClassMap()
        {
            Map(p => p.CustomIndex, ignoreCase: true);
        }
    }

    private class CustomIndexNullableEnumPropertyClass
    {
        [ExcelColumnIndex(1)]
        public CustomEnum? CustomIndex { get; set; }
    }

    private class DefaultCustomIndexNullableEnumPropertyClassMap : ExcelClassMap<CustomIndexNullableEnumPropertyClass>
    {
        public DefaultCustomIndexNullableEnumPropertyClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class CustomCustomIndexNullableEnumPropertyClassMap : ExcelClassMap<CustomIndexNullableEnumPropertyClass>
    {
        public CustomCustomIndexNullableEnumPropertyClassMap()
        {
            Map(p => p.CustomIndex, ignoreCase: true);
        }
    }
    
    private class CustomIndexEnumFieldClass
    {
        [ExcelColumnIndex(1)]
        public CustomEnum CustomIndex { get; set; }
    }

    private class DefaultCustomIndexEnumFieldClassMap : ExcelClassMap<CustomIndexEnumFieldClass>
    {
        public DefaultCustomIndexEnumFieldClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class CustomCustomIndexEnumFieldClassMap : ExcelClassMap<CustomIndexEnumFieldClass>
    {
        public CustomCustomIndexEnumFieldClassMap()
        {
            Map(p => p.CustomIndex, ignoreCase: true);
        }
    }

    private class CustomIndexNullableEnumFieldClass
    {
        [ExcelColumnIndex(1)]
        public CustomEnum? CustomIndex { get; set; }
    }

    private class DefaultCustomIndexNullableEnumFieldClassMap : ExcelClassMap<CustomIndexNullableEnumFieldClass>
    {
        public DefaultCustomIndexNullableEnumFieldClassMap()
        {
            Map(p => p.CustomIndex);
        }
    }

    private class CustomCustomIndexNullableEnumFieldClassMap : ExcelClassMap<CustomIndexNullableEnumFieldClass>
    {
        public CustomCustomIndexNullableEnumFieldClassMap()
        {
            Map(p => p.CustomIndex, ignoreCase: true);
        }
    }

    private enum CustomEnum
    {
        a,
        B
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Equal(new object?[] { "1", null, "2" }, row2.CustomValue);

        var row3 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexEnumerableMultipleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_AutoMappedCustomIndexEnumerableMultipleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Equal(new object?[] { "1", null, "2" }, row2.CustomValue);

        var row3 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexEnumerableProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumerablePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexEnumerableField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Equal(new object?[] { "1", null, "2" }, row2.CustomValue);

        var row3 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexEnumerableMultipleProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumerableMultiplePropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Equal(["1", null, "2"], row2.CustomValue);

        var row3 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomIndexEnumerableMultiplePropertyClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomIndexEnumerableMultipleField_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumerableMultipleFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Equal(["1", "2", "3"], row1.CustomValue);

        var row2 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Equal(new object?[] { "1", null, "2" }, row2.CustomValue);

        var row3 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Equal(["1"], row3.CustomValue);

        var row4 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Empty(row4.CustomValue);

        var row5 = sheet.ReadRow<CustomIndexEnumerableMultipleFieldClass>();
        Assert.Equal(["Invalid"], row5.CustomValue);
    }

    private class CustomIndexEnumerablePropertyClass
    {
        [ExcelColumnIndex(0)]
        public object?[] CustomValue { get; set; } = default!;
    }

    private class DefaultCustomIndexEnumerablePropertyClassMap : ExcelClassMap<CustomIndexEnumerablePropertyClass>
    {
        public DefaultCustomIndexEnumerablePropertyClassMap()
        {
            Map(p => p.CustomValue);
        }
    }

#pragma warning disable CS0649
    private class CustomIndexEnumerableFieldClass
    {
        [ExcelColumnIndex(0)]
        public object[] CustomValue = default!;
    }
#pragma warning restore CS0649

    private class DefaultCustomIndexEnumerableFieldClassMap : ExcelClassMap<CustomIndexEnumerableFieldClass>
    {
        public DefaultCustomIndexEnumerableFieldClassMap()
        {
            Map(p => p.CustomValue);
        }
    }

    private class CustomIndexEnumerableMultiplePropertyClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        [ExcelColumnIndex(0)]
        public object?[] CustomValue { get; set; } = default!;
    }

    private class DefaultCustomIndexEnumerableMultiplePropertyClassMap : ExcelClassMap<CustomIndexEnumerableMultiplePropertyClass>
    {
        public DefaultCustomIndexEnumerableMultiplePropertyClassMap()
        {
            Map(p => p.CustomValue);
        }
    }

#pragma warning disable CS0649
    private class CustomIndexEnumerableMultipleFieldClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        [ExcelColumnIndex(0)]
        public object[] CustomValue = default!;
    }
#pragma warning restore CS0649

    private class DefaultCustomIndexEnumerableMultipleFieldClassMap : ExcelClassMap<CustomIndexEnumerableMultipleFieldClass>
    {
        public DefaultCustomIndexEnumerableMultipleFieldClassMap()
        {
            Map(p => p.CustomValue);
        }
    }

    [Fact]
    public void ReadRows_CustomColumnIndexClassMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnIndex.xlsx");
        importer.Configuration.RegisterClassMap<CustomColumnIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var row1 = sheet.ReadRow<CustomColumnIndexDataClass>();
        Assert.Equal("Column3_A", row1.Value1);
        Assert.Equal("Column2_A", row1.Value2);
        Assert.Equal("Column1_A", row1.Value3);

        var row2 = sheet.ReadRow<CustomColumnIndexDataClass>();
        Assert.Equal("Column3_B", row2.Value1);
        Assert.Equal("Column2_B", row2.Value2);
        Assert.Equal("Column1_B", row2.Value3);
    }

    [Fact]
    public void ReadRows_InvalidColumnIndexClassMap1_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnIndex.xlsx");
        importer.Configuration.RegisterClassMap<InvalidColumnIndexClassMap1>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomColumnIndexDataClass>());
    }

    [Fact]
    public void ReadRows_InvalidColumnIndexClassMap2_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnIndex.xlsx");
        importer.Configuration.RegisterClassMap<InvalidColumnIndexClassMap2>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomColumnIndexDataClass>());
    }

    [Fact]
    public void ReadRows_InvalidColumnIndexOptionalClassMap1_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("CustomColumnIndex.xlsx");
        importer.Configuration.RegisterClassMap<InvalidColumnIndexOptionalClassMap1>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var row1 = sheet.ReadRow<CustomColumnIndexDataClass>();
        Assert.Null(row1.Value1);
        Assert.Null(row1.Value2);
        Assert.Null(row1.Value3);

        var row2 = sheet.ReadRow<CustomColumnIndexDataClass>();
        Assert.Null(row2.Value1);
        Assert.Null(row2.Value2);
        Assert.Null(row2.Value3);
    }

    [Fact]
    public void ReadRows_InvalidColumnIndexOptionalClassMap2_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("CustomColumnIndex.xlsx");
        importer.Configuration.RegisterClassMap<InvalidColumnIndexOptionalClassMap2>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var row1 = sheet.ReadRow<CustomColumnIndexDataClass>();
        Assert.Null(row1.Value1);
        Assert.Null(row1.Value2);
        Assert.Null(row1.Value3);

        var row2 = sheet.ReadRow<CustomColumnIndexDataClass>();
        Assert.Null(row2.Value1);
        Assert.Null(row2.Value2);
        Assert.Null(row2.Value3);
    }

    private class CustomColumnIndexDataClass
    {
        public string Value1 { get; set; } = default!;
        public string Value2 { get; set; } = default!;
        public string Value3 { get; set; } = default!;
    }

    private class CustomColumnIndexAttributeDataClass
    {
        [ExcelColumnIndex(2)]
        public string Value1 { get; set; } = default!;
        [ExcelColumnIndex(1)]
        public string Value2 { get; set; } = default!;
        [ExcelColumnIndex(0)]
        public string Value3 { get; set; } = default!;
    }

    private class DefaultColumnIndexAttributeClassMap : ExcelClassMap<CustomColumnIndexAttributeDataClass>
    {
        public DefaultColumnIndexAttributeClassMap()
        {
            Map(c => c.Value1);
            Map(c => c.Value2);
            Map(c => c.Value3);
        }
    }

    private class CustomColumnIndexClassMap : ExcelClassMap<CustomColumnIndexDataClass>
    {
        public CustomColumnIndexClassMap()
        {
            Map(c => c.Value1)
                .WithColumnIndex(2);

            Map(c => c.Value2)
                .WithColumnIndex(1);

            Map(c => c.Value3)
                .WithColumnIndex(0);
        }
    }

    private class InvalidColumnIndexClassMap1 : ExcelClassMap<CustomColumnIndexDataClass>
    {
        public InvalidColumnIndexClassMap1()
        {
            // ColumnIndex == FieldCount
            Map(c => c.Value1)
                .WithColumnIndex(3);
        }
    }

    private class InvalidColumnIndexClassMap2 : ExcelClassMap<CustomColumnIndexDataClass>
    {
        public InvalidColumnIndexClassMap2()
        {
            // ColumnIndex > FieldCount
            Map(c => c.Value1)
                .WithColumnIndex(4);
        }
    }

    private class InvalidColumnIndexOptionalClassMap1 : ExcelClassMap<CustomColumnIndexDataClass>
    {
        public InvalidColumnIndexOptionalClassMap1()
        {
            // ColumnIndex == FieldCount
            Map(c => c.Value1)
                .WithColumnIndex(3)
                .MakeOptional();
        }
    }

    private class InvalidColumnIndexOptionalClassMap2 : ExcelClassMap<CustomColumnIndexDataClass>
    {
        public InvalidColumnIndexOptionalClassMap2()
        {
            // ColumnIndex > FieldCount
            Map(c => c.Value1)
                .WithColumnIndex(4)
                .MakeOptional();
        }
    }
}
