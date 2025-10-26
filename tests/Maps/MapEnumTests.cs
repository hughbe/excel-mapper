using System.ComponentModel;

namespace ExcelMapper.Tests;

public class MapEnumTests
{
    [Fact]
    public void ReadRow_Enum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TestEnum>();
        Assert.Equal(TestEnum.Member, row1);

        // Different case cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TestEnum>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TestEnum>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TestEnum>());
    }

    [Fact]
    public void ReadRow_NullableEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TestEnum?>();
        Assert.Equal(TestEnum.Member, row1);

        // Different case cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TestEnum?>());

        // Empty cell value.
        var row3 = sheet.ReadRow<TestEnum?>();
        Assert.Null(row3);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TestEnum?>());
    }

    [Fact]
    public void ReadRow_AutoMappedEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Member, row1.Value);

        // Different case cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Member, row1.Value);

        // Different case cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());

        // Empty cell value.
        var row3 = sheet.ReadRow<NullableEnumClass>();
        Assert.Null(row3.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");
        importer.Configuration.RegisterClassMap<DefaultEnumClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Member, row1.Value);

        // Different case cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableCustomEnumClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Member, row1.Value);

        // Different case cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());

        // Empty cell value.
        var row3 = sheet.ReadRow<NullableEnumClass>();
        Assert.Null(row3.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");
        importer.Configuration.RegisterClassMap<CustomEnumClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Member, row1.Value);

        // Different case cell value.
        var row2 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Invalid, row2.Value);

        // Empty cell value.
        var row3 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Empty, row3.Value);

        // Invalid cell value.
        var row4 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Invalid, row4.Value);
    }

    [Fact]
    public void ReadRow_IgnoreCaseEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");
        importer.Configuration.RegisterClassMap<IgnoreCaseEnumClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Member, row1.Value);

        // Different case cell value.
        var row2 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Member, row2.Value);

        // Empty cell value.
        var row3 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Empty, row3.Value);

        // Invalid cell value.
        var row4 = sheet.ReadRow<EnumClass>();
        Assert.Equal(TestEnum.Invalid, row4.Value);
    }

    [Fact]
    public void ReadRow_NullableCustomMappedEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableCustomEnumClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Member, row1.Value);

        // Different case cell value.
        var row2 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Invalid, row2.Value);

        // Empty cell value.
        var row3 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Empty, row3.Value);

        // Invalid cell value.
        var row4 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Invalid, row4.Value);
    }

    [Fact]
    public void ReadRow_NullableIgnoreCaseEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums.xlsx");
        importer.Configuration.RegisterClassMap<IgnoreCaseNullableCustomEnumClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Member, row1.Value);

        // Different case cell value.
        var row2 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Member, row2.Value);

        // Empty cell value.
        var row3 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Empty, row3.Value);

        // Invalid cell value.
        var row4 = sheet.ReadRow<NullableEnumClass>();
        Assert.Equal(TestEnum.Invalid, row4.Value);
    }

    private enum TestEnum
    {
        Member,
        Empty,
        Invalid
    }

    private class EnumClass
    {
        public TestEnum Value { get; set; }
    }

    private class DefaultEnumClassMap : ExcelClassMap<EnumClass>
    {
        public DefaultEnumClassMap()
        {
            Map(u => u.Value);
        }
    }

    private class CustomEnumClassMap : ExcelClassMap<EnumClass>
    {
        public CustomEnumClassMap()
        {
            Map(u => u.Value)
                .WithEmptyFallback(TestEnum.Empty)
                .WithInvalidFallback(TestEnum.Invalid);
        }
    }

    private class IgnoreCaseEnumClassMap : ExcelClassMap<EnumClass>
    {
        public IgnoreCaseEnumClassMap()
        {
            Map(u => u.Value, ignoreCase: true)
                .WithEmptyFallback(TestEnum.Empty)
                .WithInvalidFallback(TestEnum.Invalid);
        }
    }

    private class NullableEnumClass
    {
        public TestEnum? Value { get; set; }
    }

    private class DefaultNullableCustomEnumClassMap : ExcelClassMap<NullableEnumClass>
    {
        public DefaultNullableCustomEnumClassMap()
        {
            Map(u => u.Value);
        }
    }

    private class CustomNullableCustomEnumClassMap : ExcelClassMap<NullableEnumClass>
    {
        public CustomNullableCustomEnumClassMap()
        {
            Map(u => u.Value)
                .WithEmptyFallback(TestEnum.Empty)
                .WithInvalidFallback(TestEnum.Invalid);
        }
    }

    private class IgnoreCaseNullableCustomEnumClassMap : ExcelClassMap<NullableEnumClass>
    {
        public IgnoreCaseNullableCustomEnumClassMap()
        {
            Map(u => u.Value, ignoreCase: true)
                .WithEmptyFallback(TestEnum.Empty)
                .WithInvalidFallback(TestEnum.Invalid);
        }
    }

    [Fact]
    public void ReadRow_DescriptionAttributeEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<DescriptionEnum>();
        Assert.Equal(DescriptionEnum.First, row1);

        var row2 = sheet.ReadRow<DescriptionEnum>();
        Assert.Equal(DescriptionEnum.First, row2);

        var row3 = sheet.ReadRow<DescriptionEnum>();
        Assert.Equal(DescriptionEnum.Second, row3);

        var row4 = sheet.ReadRow<DescriptionEnum>();
        Assert.Equal(DescriptionEnum.Third, row4);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnum>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnum>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnum>());
    }

    private enum DescriptionEnum
    {
        [Description("First Value")]
        First,

        [Description("Second Value")]
        Second,

        [Description("Third Value")]
        Third
    }

    [Fact]
    public void ReadRow_DescriptionAttributeNullableEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<DescriptionEnum?>();
        Assert.Equal(DescriptionEnum.First, row1);

        var row2 = sheet.ReadRow<DescriptionEnum?>();
        Assert.Equal(DescriptionEnum.First, row2);

        var row3 = sheet.ReadRow<DescriptionEnum?>();
        Assert.Equal(DescriptionEnum.Second, row3);

        var row4 = sheet.ReadRow<DescriptionEnum?>();
        Assert.Equal(DescriptionEnum.Third, row4);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnum?>());

        // Empty cell value.
        var row6 = sheet.ReadRow<DescriptionEnum?>();
        Assert.Null(row6);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnum>());
    }

    [Fact]
    public void ReadRow_InvalidDescriptionAttributeEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<InvalidDescriptionEnum>();
        Assert.Equal(InvalidDescriptionEnum.First, row1);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum>());

        var row4 = sheet.ReadRow<InvalidDescriptionEnum>();
        Assert.Equal(InvalidDescriptionEnum.Third, row4);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum>());
    }

    private enum InvalidDescriptionEnum
    {
        [Description(null!)]
        First,

        [Description("")]
        Second,

        [Description("Third Value")]
        Third
    }

    [Fact]
    public void ReadRow_InvalidDescriptionAttributeNullableEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<InvalidDescriptionEnum?>();
        Assert.Equal(InvalidDescriptionEnum.First, row1);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum?>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum?>());

        var row4 = sheet.ReadRow<InvalidDescriptionEnum?>();
        Assert.Equal(InvalidDescriptionEnum.Third, row4);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum?>());

        // Empty cell value.
        var row6 = sheet.ReadRow<InvalidDescriptionEnum?>();
        Assert.Null(row6);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnum?>());
    }

    [Fact]
    public void ReadRow_AutoMappedDescriptionAttributeEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<DescriptionEnumClass>();
        Assert.Equal(DescriptionEnum.First, row1.Value);

        var row2 = sheet.ReadRow<DescriptionEnumClass>();
        Assert.Equal(DescriptionEnum.First, row2.Value);

        var row3 = sheet.ReadRow<DescriptionEnumClass>();
        Assert.Equal(DescriptionEnum.Second, row3.Value);

        var row4 = sheet.ReadRow<DescriptionEnumClass>();
        Assert.Equal(DescriptionEnum.Third, row4.Value);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnumClass>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnumClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnumClass>());
    }

    private class DescriptionEnumClass
    {
        public DescriptionEnum Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDescriptionAttributeEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");
        importer.Configuration.RegisterClassMap<DescriptionEnumClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<DescriptionEnumClass>();
        Assert.Equal(DescriptionEnum.First, row1.Value);

        var row2 = sheet.ReadRow<DescriptionEnumClass>();
        Assert.Equal(DescriptionEnum.First, row2.Value);

        var row3 = sheet.ReadRow<DescriptionEnumClass>();
        Assert.Equal(DescriptionEnum.Second, row3.Value);

        var row4 = sheet.ReadRow<DescriptionEnumClass>();
        Assert.Equal(DescriptionEnum.Third, row4.Value);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnumClass>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnumClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionEnumClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedDescriptionAttributeNullableEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Equal(DescriptionEnum.First, row1.Value);

        var row2 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Equal(DescriptionEnum.First, row2.Value);

        var row3 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Equal(DescriptionEnum.Second, row3.Value);

        var row4 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Equal(DescriptionEnum.Third, row4.Value);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionNullableEnumClass>());

        // Empty cell value.
        var row6 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Null(row6.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionNullableEnumClass>());
    }

    private class DescriptionNullableEnumClass
    {
        public DescriptionEnum? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDescriptionAttributeNullableEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");
        importer.Configuration.RegisterClassMap<DescriptionNullableEnumClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Equal(DescriptionEnum.First, row1.Value);

        var row2 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Equal(DescriptionEnum.First, row2.Value);

        var row3 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Equal(DescriptionEnum.Second, row3.Value);

        var row4 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Equal(DescriptionEnum.Third, row4.Value);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionNullableEnumClass>());

        // Empty cell value.
        var row6 = sheet.ReadRow<DescriptionNullableEnumClass>();
        Assert.Null(row6.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DescriptionNullableEnumClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedDescriptionAttributeInvalidEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<InvalidDescriptionEnumClass>();
        Assert.Equal(InvalidDescriptionEnum.First, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());

        var row4 = sheet.ReadRow<InvalidDescriptionEnumClass>();
        Assert.Equal(InvalidDescriptionEnum.Third, row4.Value);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());
    }

    private class InvalidDescriptionEnumClass
    {
        public InvalidDescriptionEnum Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDescriptionAttributeInvalidEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");
        importer.Configuration.RegisterClassMap<InvalidDescriptionEnumClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<InvalidDescriptionEnumClass>();
        Assert.Equal(InvalidDescriptionEnum.First, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());

        var row4 = sheet.ReadRow<InvalidDescriptionEnumClass>();
        Assert.Equal(InvalidDescriptionEnum.Third, row4.Value);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionEnumClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedDescriptionAttributeNullableInvalidEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<InvalidDescriptionNullableEnumClass>();
        Assert.Equal(InvalidDescriptionEnum.First, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionNullableEnumClass>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionNullableEnumClass>());

        var row4 = sheet.ReadRow<InvalidDescriptionNullableEnumClass>();
        Assert.Equal(InvalidDescriptionEnum.Third, row4.Value);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionNullableEnumClass>());

        // Empty cell value.
        var row6 = sheet.ReadRow<InvalidDescriptionNullableEnumClass>();
        Assert.Null(row6.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionNullableEnumClass>());
    }

    private class InvalidDescriptionNullableEnumClass
    {
        public InvalidDescriptionEnum? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDescriptionAttributeNullableInvalidEnum_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Enums_Description.xlsx");
        importer.Configuration.RegisterClassMap<InvalidDescriptionNullableEnumClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<InvalidDescriptionNullableEnumClass>();
        Assert.Equal(InvalidDescriptionEnum.First, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionNullableEnumClass>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionNullableEnumClass>());

        var row4 = sheet.ReadRow<InvalidDescriptionNullableEnumClass>();
        Assert.Equal(InvalidDescriptionEnum.Third, row4.Value);

        // Lower case.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionNullableEnumClass>());

        // Empty cell value.
        var row6 = sheet.ReadRow<InvalidDescriptionNullableEnumClass>();
        Assert.Null(row6.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidDescriptionNullableEnumClass>());
    }
}
