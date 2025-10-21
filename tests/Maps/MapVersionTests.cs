using System;
using Xunit;

namespace ExcelMapper.Tests;

public class MapVersionTests
{
    [Fact]
    public void ReadRow_Version_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Versions.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Version>();
        Assert.Equal(new Version(1, 0), row1);

        var row2 = sheet.ReadRow<Version>();
        Assert.Equal(new Version(1, 2), row2);

        var row3 = sheet.ReadRow<Version>();
        Assert.Equal(new Version(1, 2, 3), row3);

        var row4 = sheet.ReadRow<Version>();
        Assert.Equal(new Version(1, 2, 3, 4), row4);

        // Empty cell value.
        Assert.Null(sheet.ReadRow<Version>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Version>());
    }

    [Fact]
    public void ReadRow_AutoMappedVersion_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Versions.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 0), row1.Value);

        var row2 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2), row2.Value);

        var row3 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2, 3), row3.Value);

        var row4 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2, 3, 4), row4.Value);

        // Empty cell value.
        var row5 = sheet.ReadRow<VersionClass>();
        Assert.Null(row5.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<VersionClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedVersion_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Versions.xlsx");
        importer.Configuration.RegisterClassMap<VersionClass>(c =>
        {
            c.Map(v => v.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 0), row1.Value);

        var row2 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2), row2.Value);

        var row3 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2, 3), row3.Value);

        var row4 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2, 3, 4), row4.Value);

        // Empty cell value.
        var row5 = sheet.ReadRow<VersionClass>();
        Assert.Null(row5.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<VersionClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedVersion_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Versions.xlsx");
        importer.Configuration.RegisterClassMap<VersionClass>(c =>
        {
            c.Map(v => v.Value)
                .WithEmptyFallback(new Version(0, 0))
                .WithInvalidFallback(new Version(9, 9, 9, 9));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 0), row1.Value);

        var row2 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2), row2.Value);

        var row3 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2, 3), row3.Value);

        var row4 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(1, 2, 3, 4), row4.Value);

        // Empty cell value.
        var row5 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(0, 0), row5.Value);

        // Invalid cell value.
        var row6 = sheet.ReadRow<VersionClass>();
        Assert.Equal(new Version(9, 9, 9, 9), row6.Value);
    }

    private class VersionClass
    {
        public Version Value { get; set; } = default!;
    }
}
