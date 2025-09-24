using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
    public class AllColumnNamesValueReaderTests
    {
        

        [Fact]
        public void TryGetValues_InvokeCanRead_Success()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            var reader = new AllColumnNamesValueReader();
            IEnumerable<ReadCellValueResult>? result = null;
            Assert.True(reader.TryGetValues(sheet, 0, importer.Reader, out result));
            Assert.Equal(["Value"], result.Select(r => r.StringValue));
        }

        [Fact]
        public void TryGetValues_InvokeNullSheet_ThrowsArgumentNullException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            var reader = new AllColumnNamesValueReader();
            IEnumerable<ReadCellValueResult>? result = null;
            Assert.Throws<ArgumentNullException>("sheet", () => reader.TryGetValues(null!, 0, importer.Reader, out result));
            Assert.Null(result);
        }

        [Fact]
        public void TryGetValues_InvokeSheetWithoutHeadingHasHeading_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();

            var reader = new AllColumnNamesValueReader();
            IEnumerable<ReadCellValueResult>? result = null;
            Assert.Throws<ExcelMappingException>(() => reader.TryGetValues(sheet, 0, importer.Reader, out result));
            Assert.Null(result);
        }

        [Fact]
        public void TryGetValues_InvokeSheetWithoutHeadingHasNoHeading_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            var reader = new AllColumnNamesValueReader();
            IEnumerable<ReadCellValueResult>? result = null;
            Assert.Throws<ExcelMappingException>(() => reader.TryGetValues(sheet, 0, importer.Reader, out result));
            Assert.Null(result);
        }
    }
}
