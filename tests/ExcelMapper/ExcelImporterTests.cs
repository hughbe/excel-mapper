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
        public void Ctor_Stream()
        {
            using (var stream = Helpers.GetResource("Primitives.xlsx"))
            using (var importer = new ExcelImporter(stream))
            {
                Assert.Equal("Primitives", importer.ReadSheet().Name);
                Assert.Equal(3, importer.NumberOfSheets);
            }
        }

        [Fact]
        public void Ctor_NullStream_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("stream", () => new ExcelImporter((Stream)null));
        }

        [Fact]
        public void Ctor_IExcelDataReader()
        {
            using (var stream = Helpers.GetResource("Primitives.xlsx"))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            using (var importer = new ExcelImporter(reader))
            {
                Assert.Equal("Primitives", importer.ReadSheet().Name);
                Assert.Equal(3, importer.NumberOfSheets);
            }
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
        public void ReadSheets_InvokeReadSheetsAfter_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet[] sheets = importer.ReadSheets().ToArray();
                Assert.Equal(new string[] { "Primitives", "Empty", "Third Sheet" }, sheets.Select(sheet => sheet.Name));
                Assert.Equal(new int[] { 0, 1, 2 }, sheets.Select(sheet => sheet.Index));
                Assert.Equal(new bool[] { true, true, true }, sheets.Select(sheet => sheet.HasHeading));
                Assert.Equal(new ExcelHeading[] { null, null, null }, sheets.Select(sheet => sheet.Heading));

                sheets = importer.ReadSheets().ToArray();
                Assert.Equal(new string[] { "Primitives", "Empty", "Third Sheet" }, sheets.Select(sheet => sheet.Name));
                Assert.Equal(new int[] { 0, 1, 2 }, sheets.Select(sheet => sheet.Index));
                Assert.Equal(new bool[] { true, true, true }, sheets.Select(sheet => sheet.HasHeading));
                Assert.Equal(new ExcelHeading[] { null, null, null }, sheets.Select(sheet => sheet.Heading));
            }
        }

        [Fact]
        public void ReadSheets_InvokeReadSheetAfter_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet[] sheets = importer.ReadSheets().ToArray();
                Assert.Equal(new string[] { "Primitives", "Empty", "Third Sheet" }, sheets.Select(sheet => sheet.Name));
                Assert.Equal(new int[] { 0, 1, 2 }, sheets.Select(sheet => sheet.Index));
                Assert.Equal(new bool[] { true, true, true }, sheets.Select(sheet => sheet.HasHeading));
                Assert.Equal(new ExcelHeading[] { null, null, null }, sheets.Select(sheet => sheet.Heading));

                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
            }
        }

        [Fact]
        public void ReadSheet_AllSheets_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                Assert.Equal("Primitives", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(0, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                sheet = importer.ReadSheet();
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                sheet = importer.ReadSheet();
                Assert.Equal("Third Sheet", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(2, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
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

                Assert.Throws<ExcelMappingException>(() => importer.ReadSheet());
            }
        }

        [Fact]
        public void TryReadSheet_AllSheets_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.Equal("Primitives", sheet1.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet1.Visibility);
                Assert.Equal(0, sheet1.Index);
                Assert.True(sheet1.HasHeading);
                Assert.Null(sheet1.Heading);
                Assert.Equal(-1, sheet1.CurrentRowIndex);

                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.Equal("Empty", sheet2.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet2.Visibility);
                Assert.Equal(1, sheet2.Index);
                Assert.True(sheet2.HasHeading);
                Assert.Null(sheet2.Heading);
                Assert.Equal(-1, sheet2.CurrentRowIndex);

                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));
                Assert.Equal("Third Sheet", sheet3.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet3.Visibility);
                Assert.Equal(2, sheet3.Index);
                Assert.True(sheet3.HasHeading);
                Assert.Null(sheet3.Heading);
                Assert.Equal(-1, sheet3.CurrentRowIndex);
            }
        }

        [Fact]
        public void TryReadSheet_NoMoreSheets_ReturnsFalse()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.ReadSheet();
                importer.ReadSheet();
                importer.ReadSheet();

                Assert.False(importer.TryReadSheet(out ExcelSheet sheet));
                Assert.Null(sheet);
            }
        }

        [Fact]
        public void ReadSheet_SheetNameExistsNotAlreadyRead_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet("Empty");
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                // Reading a named sheet should reset the reader after finding the column.
                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
            }
        }

        [Fact]
        public void ReadSheet_SheetNameExistsAlreadyRead_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.ReadSheet();
                importer.ReadSheet();
                importer.ReadSheet();

                // Reading a named sheet should reset the reader before finding the sheet.
                ExcelSheet sheet = importer.ReadSheet("Empty");
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                // Reading a named sheet should reset the reader after finding the column.
                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
            }
        }

        [Fact]
        public void ReadSheet_NullSheetName_ThrowsArgumentNullException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.Throws<ArgumentNullException>("sheetName", () => importer.ReadSheet(null));
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("empty")]
        [InlineData(" Empty ")]
        [InlineData("invalid")]
        [InlineData("  \r \t  ")]
        public void ReadSheet_NoSuchSheet_ThrowsExcelMappingException(string sheetName)
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.Throws<ExcelMappingException>(() => importer.ReadSheet(sheetName));
            }
        }

        [Fact]
        public void TryReadSheet_SheetNameExistsNotAlreadyRead_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet("Empty", out ExcelSheet sheet));
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                // Reading a named sheet should reset the reader after finding the column.
                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
            }
        }

        [Fact]
        public void TryReadSheet_SheetNameExistsAlreadyRead_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.ReadSheet();
                importer.ReadSheet();
                importer.ReadSheet();

                // Reading a named sheet should reset the reader before finding the sheet.
                Assert.True(importer.TryReadSheet("Empty", out ExcelSheet sheet));
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                // Reading a named sheet should reset the reader after finding the column.
                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
            }
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("empty")]
        [InlineData(" Empty ")]
        [InlineData("invalid")]
        [InlineData("  \r \t  ")]
        public void TryReadSheet_NoSuchSheet_ReturnsFalse(string sheetName)
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.False(importer.TryReadSheet(sheetName, out ExcelSheet sheet));
                Assert.Null(sheet);
            }
        }

        [Fact]
        public void ReadSheet_SheetIndexExistsNotAlreadyRead_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet(1);
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                // Reading a named sheet should reset the reader after finding the column.
                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
            }
        }

        [Fact]
        public void ReadSheet_SheetIndexExistsAlreadyRead_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.ReadSheet();
                importer.ReadSheet();
                importer.ReadSheet();

                // Reading a named sheet should reset the reader before finding the sheet.
                ExcelSheet sheet = importer.ReadSheet(1);
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                // Reading a named sheet should reset the reader after finding the column.
                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
            }
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(3)]
        public void ReadSheet_InvalidSheetIndex_ThrowsArgumentOutOfRangeException(int sheetIndex)
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.Throws<ArgumentOutOfRangeException>("sheetIndex", () => importer.ReadSheet(sheetIndex));
            }
        }

        [Fact]
        public void TryReadSheet_SheetIndexExistsNotAlreadyRead_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(1, out ExcelSheet sheet));
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                // Reading a named sheet should reset the reader after finding the column.
                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
            }
        }

        [Fact]
        public void TryReadSheet_SheetIndexExistsAlreadyRead_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.ReadSheet();
                importer.ReadSheet();
                importer.ReadSheet();

                // Reading a named sheet should reset the reader before finding the sheet.
                Assert.True(importer.TryReadSheet(1, out ExcelSheet sheet));
                Assert.Equal("Empty", sheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
                Assert.Equal(1, sheet.Index);
                Assert.True(sheet.HasHeading);
                Assert.Null(sheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);

                // Reading a named sheet should reset the reader after finding the column.
                ExcelSheet nextSheet = importer.ReadSheet();
                Assert.Equal("Primitives", nextSheet.Name);
                Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
                Assert.Equal(0, nextSheet.Index);
                Assert.True(nextSheet.HasHeading);
                Assert.Null(nextSheet.Heading);
                Assert.Equal(-1, sheet.CurrentRowIndex);
            }
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(3)]
        public void TryReadSheet_InvalidSheetIndex_ReturnsFalse(int sheetIndex)
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.False(importer.TryReadSheet(sheetIndex, out ExcelSheet sheet));
                Assert.Null(sheet);
            }
        }

        [Fact]
        public void MultipleSheets_ReadFirstHeadingThenFirstRowsThenLastHeadingThenLastRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.ReadHeading();

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);

                sheet3.ReadHeading();

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadFirstHeadingThenLastHeadingThenFirstRowsThenLastRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.ReadHeading();
                sheet3.ReadHeading();

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadFirstHeadingThenLastHeadingThenLastRowsThenFirstRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.ReadHeading();
                sheet3.ReadHeading();

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadLastHeadingThenLastRowsThenFirstHeadingThenFirstRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet3.ReadHeading();

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);

                sheet1.ReadHeading();

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadLastHeadingThenFirstHeadingThenFirstRowsThenLastRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet3.ReadHeading();
                sheet1.ReadHeading();

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadLastHeadingThenFirstHeadingThenLastRowsThenFirstRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet3.ReadHeading();
                sheet1.ReadHeading();

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadFirstRowThenLastRowThenFirstRowThenLastRow_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.ReadHeading();
                sheet3.ReadHeading();

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadFirstRowThenLastRowThenLastRowThenFirstRow_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.ReadHeading();
                sheet3.ReadHeading();

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadLastRowThenFirstRowThenLastRowThenFirstRow_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.ReadHeading();
                sheet3.ReadHeading();

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_ReadLastRowThenFirstRowThenLastRowThenLastRowThenFirstRow_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.ReadHeading();
                sheet3.ReadHeading();

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row1.StringValue);

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row2.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s3", sheet3Row3.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Null(sheet1Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_NoHeadingReadFirstRowsThenLastRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.HasHeading = false;
                sheet3.HasHeading = false;

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row3.StringValue);

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_NoHeadingReadLastRowsThenFirstRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.HasHeading = false;
                sheet3.HasHeading = false;

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row3.StringValue);

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_NoHeadingReadFirstRowsThenLastRowsThenFirstRowsThenLastRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.HasHeading = false;
                sheet3.HasHeading = false;

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet1Row1.StringValue);

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet3Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row2.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row3.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_NoHeadingReadFirstRowsThenLastRowsThenLastRowsThenFirstRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.HasHeading = false;
                sheet3.HasHeading = false;

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet1Row1.StringValue);

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet3Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row2.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row2.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row3.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_NoHeadingReadLastRowsThenFirstRowsThenFirstRowsThenLastRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.HasHeading = false;
                sheet3.HasHeading = false;

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet3Row1.StringValue);

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet1Row1.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row2.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row3.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row3.StringValue);
            }
        }

        [Fact]
        public void MultipleSheets_NoHeadingReadLastRowsThenFirstRowsThenLastRowsThenFirstRows_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

                Assert.True(importer.TryReadSheet(out ExcelSheet sheet1));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet2));
                Assert.True(importer.TryReadSheet(out ExcelSheet sheet3));

                sheet1.HasHeading = false;
                sheet3.HasHeading = false;

                TestClass sheet3Row1 = sheet3.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet3Row1.StringValue);

                TestClass sheet1Row1 = sheet1.ReadRow<TestClass>();
                Assert.Equal("StringValue", sheet1Row1.StringValue);

                TestClass sheet3Row2 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s1", sheet3Row2.StringValue);
    
                TestClass sheet1Row2 = sheet1.ReadRow<TestClass>();
                Assert.Equal("a", sheet1Row2.StringValue);

                TestClass sheet3Row3 = sheet3.ReadRow<TestClass>();
                Assert.Equal("s2", sheet3Row3.StringValue);

                TestClass sheet1Row3 = sheet1.ReadRow<TestClass>();
                Assert.Equal("b", sheet1Row3.StringValue);
            }
        }

        private class TestClassMapColumnIndex : ExcelClassMap<TestClass>
        {
            public TestClassMapColumnIndex()
            {
                Map(value => value.StringValue)
                    .WithColumnIndex(1);
            }
        }

        private class TestClass
        {
            public string StringValue { get; set; }
        }
    }
}
