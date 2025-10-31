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

        // Relative cell value.
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

        // Relative cell value.
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

    [Fact]
    public void ReadRow_CustomMappedUriKindRelativeOrAbsolute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<CustomUriClass>(c =>
        {
            c.Map(u => u.Value)
                .WithUriKind(UriKind.RelativeOrAbsolute)
                .WithColumnName("Uri")
                .WithEmptyFallback(new Uri("http://empty.com/"))
                .WithInvalidFallback(new Uri("http://invalid.com/"));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Value);

        var row2 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("http://empty.com"), row2.Value);

        var row3 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("invalid", UriKind.Relative), row3.Value);
    }

    private class CustomUriClass
    {
        public Uri? Value { get; set; }
    }

    [Fact]
    public void ReadRow_CustomMappedUriKindAbsolute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<CustomUriClass>(c =>
        {
            c.Map(u => u.Value)
                .WithUriKind(UriKind.Absolute)
                .WithColumnName("Uri")
                .WithEmptyFallback(new Uri("http://empty.com/"))
                .WithInvalidFallback(new Uri("http://invalid.com/"));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Value);

        var row2 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("http://empty.com"), row2.Value);

        var row3 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("http://invalid.com"), row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedUriKindRelative_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<CustomUriClass>(c =>
        {
            c.Map(u => u.Value)
                .WithUriKind(UriKind.Relative)
                .WithColumnName("Uri")
                .WithEmptyFallback(new Uri("http://empty.com/"))
                .WithInvalidFallback(new Uri("http://invalid.com/"));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("http://invalid.com"), row1.Value);

        var row2 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("http://empty.com"), row2.Value);

        var row3 = sheet.ReadRow<CustomUriClass>();
        Assert.Equal(new Uri("invalid", UriKind.Relative), row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedUriKindAttributeAbsolute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Absolute cell value.
        var row1 = sheet.ReadRow<UriKindAbsoluteClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        // Empty cell value.
        var row2 = sheet.ReadRow<UriKindAbsoluteClass>();
        Assert.Null(row2.Uri);

        // Relative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriKindAbsoluteClass>());
    }

    private class UriKindAbsoluteClass
    {
        [ExcelUri(UriKind.Absolute)]
        public Uri Uri { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedUriKindAttributeAbsolute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<UriKindAbsoluteClass>(c =>
        {
            c.Map(u => u.Uri);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Absolute cell value.
        var row1 = sheet.ReadRow<UriKindAbsoluteClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        // Empty cell value.
        var row2 = sheet.ReadRow<UriKindAbsoluteClass>();
        Assert.Null(row2.Uri);

        // Relative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriKindAbsoluteClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedUriKindAttributeRelative_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Absolute cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriKindRelativeClass>());

        // Empty cell value.
        var row2 = sheet.ReadRow<UriKindRelativeClass>();
        Assert.Null(row2.Uri);

        // Relative cell value.
        var row3 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Equal(new Uri("invalid", UriKind.Relative), row3.Uri);
    }

    private class UriKindRelativeClass
    {
        [ExcelUri(UriKind.Relative)]
        public Uri Uri { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedUriKindAttributeRelative_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<UriKindRelativeClass>(c =>
        {
            c.Map(u => u.Uri);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Absolute cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriKindRelativeClass>());

        // Empty cell value.
        var row2 = sheet.ReadRow<UriKindRelativeClass>();
        Assert.Null(row2.Uri);

        // Relative cell value.
        var row3 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Equal(new Uri("invalid", UriKind.Relative), row3.Uri);
    }


    [Fact]
    public void ReadRow_AutoMappedUriKindAttributeRelativeOrAbsolute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Absolute cell value.
        var row1 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        // Empty cell value.
        var row2 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Null(row2.Uri);

        // Relative cell value.
        var row3 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Equal(new Uri("invalid", UriKind.Relative), row3.Uri);
    }

    private class UriKindRelativeOrAbsoluteClass
    {
        [ExcelUri(UriKind.RelativeOrAbsolute)]
        public Uri Uri { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedUriKindAttributeRelativeOrAbsolute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<UriKindRelativeOrAbsoluteClass>(c =>
        {
            c.Map(u => u.Uri);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Absolute cell value.
        var row1 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        // Empty cell value.
        var row2 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Null(row2.Uri);

        // Relative cell value.
        var row3 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Equal(new Uri("invalid", UriKind.Relative), row3.Uri);
    }
}
