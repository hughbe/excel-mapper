using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelMappingExceptionTests
    {
        [Fact]
        public void Ctor_Default()
        {
            var exception = new ExcelMappingException();
            Assert.NotNull(exception.Message);
            Assert.Null(exception.InnerException);
        }

        [Fact]
        public void Ctor_Message()
        {
            var exception = new ExcelMappingException("message");
            Assert.Equal("message", exception.Message);
            Assert.Null(exception.InnerException);
        }

        [Fact]
        public void Ctor_Message_InnerException()
        {
            var innerException = new DivideByZeroException();
            var exception = new ExcelMappingException("message", innerException);
            Assert.Equal("message", exception.Message);
            Assert.Same(innerException, exception.InnerException);
        }

        [Fact]
        public void Ctor_Message_SheetWithReadHeading_RowIndex_ColumnIndex()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                var exception = new ExcelMappingException("Message", sheet, 10, 1);
                Assert.Equal("Message in column \"StringValue\" on row 10 in sheet \"Primitives\".", exception.Message);
                Assert.Null(exception.InnerException);
            }
        }

        [Fact]
        public void Ctor_Message_SheetWithNonReadHeading_RowIndex_ColumnIndex()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                var exception = new ExcelMappingException("Message", sheet, 10, 1);
                Assert.Equal("Message in column \"StringValue\" on row 10 in sheet \"Primitives\".", exception.Message);
                Assert.Null(exception.InnerException);
            }
        }
    }
}
