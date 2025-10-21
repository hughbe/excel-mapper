using System;
using System.Numerics;
using Xunit;

namespace ExcelMapper.Tests;

public class MapBigIntegerTests
{
    [Fact]
    public void ReadRow_BigInteger_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<BigInteger>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigInteger>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigInteger>());
    }

    [Fact]
    public void ReadRow_NullableBigInteger_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<BigInteger?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<BigInteger?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigIntegerValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedBigInteger_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<BigIntegerValue>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigIntegerValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigIntegerValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableBigInteger_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableBigIntegerClass>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableBigIntegerClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigIntegerValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedBigInteger_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultBigIntegerValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<BigIntegerValue>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigIntegerValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigIntegerValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableBigInteger_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableBigIntegerClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableBigIntegerClass>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableBigIntegerClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigIntegerValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedBigInteger_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomBigIntegerValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<BigIntegerValue>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<BigIntegerValue>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<BigIntegerValue>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableBigInteger_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableBigIntegerClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableBigIntegerClass>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableBigIntegerClass>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableBigIntegerClass>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_BigIntegerOverflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigInteger>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BigInteger>());
    }

    private class BigIntegerValue
    {
        public BigInteger Value { get; set; }
    }

    private class DefaultBigIntegerValueMap : ExcelClassMap<BigIntegerValue>
    {
        public DefaultBigIntegerValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomBigIntegerValueMap : ExcelClassMap<BigIntegerValue>
    {
        public CustomBigIntegerValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }

    private class NullableBigIntegerClass
    {
        public BigInteger? Value { get; set; }
    }

    private class DefaultNullableBigIntegerClassMap : ExcelClassMap<NullableBigIntegerClass>
    {
        public DefaultNullableBigIntegerClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableBigIntegerClassMap : ExcelClassMap<NullableBigIntegerClass>
    {
        public CustomNullableBigIntegerClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback((BigInteger)11)
                .WithInvalidFallback((BigInteger)10);
        }
    }
}
