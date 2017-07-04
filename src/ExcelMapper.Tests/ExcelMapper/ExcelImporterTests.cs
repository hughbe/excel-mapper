using System;
using System.IO;
using System.Linq;
using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelImporterTests
    {
        [Fact]
        public void Ctor_NullStream_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("stream", () => new ExcelImporter((Stream)null));
        }

        [Fact]
        public void Ctor_NullReader_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("reader", () => new ExcelImporter((IExcelDataReader)null));
        }

        [Fact]
        public void ReadSheets_Invoke_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet[] sheets = importer.ReadSheets().ToArray();
                Assert.Equal(new string[] { "Primitives", "Empty", "Third Sheet" }, sheets.Select(sheet => sheet.Name));
                Assert.Equal(new bool[] { true, true, true }, sheets.Select(sheet => sheet.HasHeading));
                Assert.Equal(new ExcelHeading[] { null, null, null }, sheets.Select(sheet => sheet.Heading));
            }
        }

        [Fact]
        public void ReadSheet_AllSheets_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                Assert.Equal("Primitives", sheet.Name);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);

                Assert.True(importer.TryReadSheet(out sheet));
                Assert.Equal("Empty", sheet.Name);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);

                Assert.True(importer.TryReadSheet(out sheet));
                Assert.Equal("Third Sheet", sheet.Name);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
            }
        }

        [Fact]
        public void ReadSheet_NoMoreSheets_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.ReadSheet();
                importer.ReadSheet();
                importer.ReadSheet();

                Assert.False(importer.TryReadSheet(out ExcelSheet sheet));
                Assert.Null(sheet);

                Assert.Throws<ExcelMappingException>(() => importer.ReadSheet());
            }
        }
    }
}
