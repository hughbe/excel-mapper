using Xunit;

namespace ExcelMapper.Tests;

public class MapSByteTests
{
    [Fact]
    public void ReadRow_SByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<sbyte>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte>());
    }

    [Fact]
    public void ReadRow_NullableSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<sbyte?>();
        Assert.Equal((sbyte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<sbyte?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<SByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableSByteClass>();
        Assert.Equal((sbyte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableSByteClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultSByteValueMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<SByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableSByteClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableSByteClass>();
        Assert.Equal((sbyte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableSByteClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomSByteValueMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<SByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<SByteValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<SByteValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableSByteClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableSByteClass>();
        Assert.Equal((sbyte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableSByteClass>();
        Assert.Equal((sbyte)-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableSByteClass>();
        Assert.Equal((sbyte)10, row3.Value);
    }

    private class SByteValue
    {
        public sbyte Value { get; set; }
    }

    private class DefaultSByteValueMap : ExcelClassMap<SByteValue>
    {
        public DefaultSByteValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomSByteValueMap : ExcelClassMap<SByteValue>
    {
        public CustomSByteValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }

    private class NullableSByteClass
    {
        public sbyte? Value { get; set; }
    }

    private class DefaultNullableSByteClassMap : ExcelClassMap<NullableSByteClass>
    {
        public DefaultNullableSByteClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableSByteClassMap : ExcelClassMap<NullableSByteClass>
    {
        public CustomNullableSByteClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }
}
