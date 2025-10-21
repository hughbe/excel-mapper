using Xunit;

namespace ExcelMapper.Tests;

public class MapDoubleTests
{
    [Fact]
    public void ReadRow_Double_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<double>();
        Assert.Equal(2.2345, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double>());
    }

    [Fact]
    public void ReadRow_NullableDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<double?>();
        Assert.Equal(2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<double?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double?>());
    }

    [Fact]
    public void ReadRow_AutoMappedDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DoubleClass>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleClass>());
    }

    private class DoubleClass
    {
        public double Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<DoubleClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DoubleClass>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<DoubleClass>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DoubleClass>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DoubleClass>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DoubleClass>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDoubleClass>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDoubleClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDoubleClass>());
    }

    private class NullableDoubleClass
    {
        public double? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableDoubleClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDoubleClass>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDoubleClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDoubleClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableDoubleClass>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDoubleClass>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDoubleClass>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDoubleClass>();
        Assert.Equal(10, row3.Value);
    }
}
