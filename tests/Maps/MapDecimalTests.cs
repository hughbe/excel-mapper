using Xunit;

namespace ExcelMapper.Tests;

public class MapDecimalTests
{
    [Fact]
    public void ReadRow_Decimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<decimal>();
        Assert.Equal(2.2345m, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal>());
    }

    [Fact]
    public void ReadRow_NullableDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<decimal?>();
        Assert.Equal(2.2345m, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<decimal?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal?>());
    }

    [Fact]
    public void ReadRow_AutoMappedDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DecimalClass>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());
    }

    private class DecimalClass
    {
        public decimal Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<DecimalClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DecimalClass>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<DecimalClass>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DecimalClass>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DecimalClass>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DecimalClass>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDecimalClass>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDecimalClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDecimalClass>());
    }

    private class NullableDecimalClass
    {
        public decimal? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableDecimalClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDecimalClass>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDecimalClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDecimalClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableDecimalClass>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDecimalClass>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDecimalClass>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDecimalClass>();
        Assert.Equal(10, row3.Value);
    }
}
