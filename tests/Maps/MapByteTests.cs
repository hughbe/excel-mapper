using Xunit;

namespace ExcelMapper.Tests;

public class MapByteTests
{
    [Fact]
    public void ReadRow_Byte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<byte>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());
    }

    [Fact]
    public void ReadRow_NullableByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<byte?>();
        Assert.Equal((byte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<byte?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableByteClass>();
        Assert.Equal((byte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableByteClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultByteValueMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableByteClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableByteClass>();
        Assert.Equal((byte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableByteClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomByteValueMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ByteValue>();
        Assert.Equal(11, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ByteValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableByteClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableByteClass>();
        Assert.Equal((byte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableByteClass>();
        Assert.Equal((byte)11, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableByteClass>();
        Assert.Equal((byte)10, row3.Value);
    }

    [Fact]
    public void ReadRow_ByteOverflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());
    }

    private class ByteValue
    {
        public byte Value { get; set; }
    }

    private class DefaultByteValueMap : ExcelClassMap<ByteValue>
    {
        public DefaultByteValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomByteValueMap : ExcelClassMap<ByteValue>
    {
        public CustomByteValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }

    private class NullableByteClass
    {
        public byte? Value { get; set; }
    }

    private class DefaultNullableByteClassMap : ExcelClassMap<NullableByteClass>
    {
        public DefaultNullableByteClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableByteClassMap : ExcelClassMap<NullableByteClass>
    {
        public CustomNullableByteClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }
}
