using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapUriTests
    {
        [Fact]
        public void ReadRow_AutoMappedUri_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Uris.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                UriValue row1 = sheet.ReadRow<UriValue>();
                Assert.Equal(new Uri("http://google.com"), row1.Uri);

                // Defaults to null if empty.
                UriValue row2 = sheet.ReadRow<UriValue>();
                Assert.Null(row2.Uri);

                // Defaults to throw if invalid.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriValue>());
            }
        }

        [Fact]
        public void ReadRow_UriWithCustomFallback_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Uris.xlsx"))
            {
                importer.Configuration.RegisterClassMap<UriValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                UriValue row1 = sheet.ReadRow<UriValue>();
                Assert.Equal(new Uri("http://google.com"), row1.Uri);

                UriValue row2 = sheet.ReadRow<UriValue>();
                Assert.Equal(new Uri("http://empty.com"), row2.Uri);

                UriValue row3 = sheet.ReadRow<UriValue>();
                Assert.Equal(new Uri("http://invalid.com"), row3.Uri);
            }
        }

        private class UriValue
        {
            public Uri Uri { get; set; }
        }

        private class UriValueFallbackMap : ExcelClassMap<UriValue>
        {
            public UriValueFallbackMap()
            {
                Map(u => u.Uri)
                    .WithEmptyFallback(new Uri("http://empty.com/"))
                    .WithInvalidFallback(new Uri("http://invalid.com/"));
            }
        }
    }
}
