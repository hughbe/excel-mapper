using Xunit;

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
}
