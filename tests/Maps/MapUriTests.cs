namespace ExcelMapper.Tests;

public class MapUriTests
{
    [Fact]
    public void ReadRow_Uri_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Uri>();
        Assert.Equal(new Uri("http://google.com"), row1);

        // Defaults to null if empty.
        var row2 = sheet.ReadRow<Uri>();
        Assert.Null(row2);

        // Defaults to throw if invalid.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Uri>());
    }

    [Fact]
    public void ReadRow_AutoMappedUri_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UriClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        // Empty cell value.
        var row2 = sheet.ReadRow<UriClass>();
        Assert.Null(row2.Uri);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriClass>());
    }

    private class UriClass
    {
        public Uri Uri { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedUri_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<UriClass>(c =>
        {
            c.Map(u => u.Uri);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UriClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        // Empty cell value.
        var row2 = sheet.ReadRow<UriClass>();
        Assert.Null(row2.Uri);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedUri_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<UriClass>(c =>
        {
            c.Map(u => u.Uri)
                .WithEmptyFallback(new Uri("http://empty.com/"))
                .WithInvalidFallback(new Uri("http://invalid.com/"));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UriClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        var row2 = sheet.ReadRow<UriClass>();
        Assert.Equal(new Uri("http://empty.com"), row2.Uri);

        var row3 = sheet.ReadRow<UriClass>();
        Assert.Equal(new Uri("http://invalid.com"), row3.Uri);
    }
}
