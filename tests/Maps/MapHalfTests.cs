using System;
using Xunit;

namespace ExcelMapper.Tests;

public class MapHalfTests
{
    [Fact]
    public void ReadRow_Half_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Half>();
        Assert.Equal((Half)2.2345, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half>());
    }

    [Fact]
    public void ReadRow_NullableHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Half?>();
        Assert.Equal((Half)2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<Half?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half?>());
    }

    [Fact]
    public void ReadRow_AutoMappedHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<HalfClass>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HalfClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HalfClass>());
    }

    private class HalfClass
    {
        public Half Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<HalfClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<HalfClass>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HalfClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HalfClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<HalfClass>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<HalfClass>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<HalfClass>();
        Assert.Equal((Half)(-10), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<HalfClass>();
        Assert.Equal((Half)10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableHalfClass>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableHalfClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableHalfClass>());
    }

    private class NullableHalfClass
    {
        public Half? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableHalfClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableHalfClass>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableHalfClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableHalfClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableHalfClass>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableHalfClass>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableHalfClass>();
        Assert.Equal((Half)(-10), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableHalfClass>();
        Assert.Equal((Half)10, row3.Value);
    }
}
