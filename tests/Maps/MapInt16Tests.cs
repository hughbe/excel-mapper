using Xunit;

namespace ExcelMapper.Tests;

public class MapInt16Tests
{
    [Fact]
    public void ReadRow_Int16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<short>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short>());
    }

    [Fact]
    public void ReadRow_NullableInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<short?>();
        Assert.Equal((short)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<short?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt16Class>();
        Assert.Equal((short)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt16Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultInt16ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableInt16ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt16Class>();
        Assert.Equal((short)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt16Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomInt16ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int16Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<Int16Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableInt16ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt16Class>();
        Assert.Equal((short)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt16Class>();
        Assert.Equal((short)-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableInt16Class>();
        Assert.Equal((short)10, row3.Value);
    }

    [Fact]
    public void ReadRow_Int16Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short>());
    }

    private class Int16Value
    {
        public short Value { get; set; }
    }

    private class DefaultInt16ValueMap : ExcelClassMap<Int16Value>
    {
        public DefaultInt16ValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomInt16ValueMap : ExcelClassMap<Int16Value>
    {
        public CustomInt16ValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }

    private class NullableInt16Class
    {
        public short? Value { get; set; }
    }

    private class DefaultNullableInt16ClassMap : ExcelClassMap<NullableInt16Class>
    {
        public DefaultNullableInt16ClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableInt16ClassMap : ExcelClassMap<NullableInt16Class>
    {
        public CustomNullableInt16ClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }
}
