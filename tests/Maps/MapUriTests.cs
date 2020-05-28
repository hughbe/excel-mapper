using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapUriTests
    {
        [Fact]
        public void ReadRow_AutoMappedUri_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Uris.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            UriClass row1 = sheet.ReadRow<UriClass>();
            Assert.Equal(new Uri("http://google.com"), row1.Uri);

            // Defaults to null if empty.
            UriClass row2 = sheet.ReadRow<UriClass>();
            Assert.Null(row2.Uri);

            // Defaults to throw if invalid.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedUri_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Uris.xlsx");
            importer.Configuration.RegisterClassMap<DefaultUriClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            UriClass row1 = sheet.ReadRow<UriClass>();
            Assert.Equal(new Uri("http://google.com"), row1.Uri);

            // Defaults to null if empty.
            UriClass row2 = sheet.ReadRow<UriClass>();
            Assert.Null(row2.Uri);

            // Defaults to throw if invalid.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriClass>());
        }

        [Fact]
        public void ReadRow_UriWithCustomFallback_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Uris.xlsx");
            importer.Configuration.RegisterClassMap<CustomUriClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            UriClass row1 = sheet.ReadRow<UriClass>();
            Assert.Equal(new Uri("http://google.com"), row1.Uri);

            UriClass row2 = sheet.ReadRow<UriClass>();
            Assert.Equal(new Uri("http://empty.com"), row2.Uri);

            UriClass row3 = sheet.ReadRow<UriClass>();
            Assert.Equal(new Uri("http://invalid.com"), row3.Uri);
        }

        private class UriClass
        {
            public Uri Uri { get; set; }
        }

        private class DefaultUriClassMap : ExcelClassMap<UriClass>
        {
            public DefaultUriClassMap()
            {
                Map(u => u.Uri);
            }
        }

        private class CustomUriClassMap : ExcelClassMap<UriClass>
        {
            public CustomUriClassMap()
            {
                Map(u => u.Uri)
                    .WithEmptyFallback(new Uri("http://empty.com/"))
                    .WithInvalidFallback(new Uri("http://invalid.com/"));
            }
        }
    }
}
