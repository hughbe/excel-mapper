using System;
using System.Collections.Generic;
using ExcelMapper.Pipeline;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelSheetTests
    {
        [Fact]
        public void ReadHeading_HasHeading_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                ExcelHeading heading = sheet.ReadHeading();
                Assert.Same(heading, sheet.Heading);

                Assert.Equal(new string[] { "Int Value", "StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "SplitValue1", "SplitValue2", "SplitValue3", "SplitValue4", "SplitValue5" }, heading.ColumnNames);
            }
        }

        [Fact]
        public void ReadHeading_AlreadyReadHeading_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
            }
        }

        [Fact]
        public void ReadHeading_DoesNotHaveHasHeading_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.HasHeading = _ => false;
                ExcelSheet sheet = importer.ReadSheet();

                Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
            }
        }

        [Fact]
        public void ReadRow_NoSuchMapping_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<PrimitiveSheet1>());
            }
        }

        [Fact]
        public void ReadRow_Sheets_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterMapping<PrimitiveSheet1Mapping>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                PrimitiveSheet1 row1 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal(1, row1.IntValue);
                Assert.Equal("a", row1.StringValue);
                Assert.True(row1.BoolValue);
                Assert.Equal(PrimitiveSheet1Enum.Government, row1.EnumValue);
                Assert.Equal(new DateTime(2017, 07, 04), row1.DateValue);
                Assert.Equal(new string[] { "a", "b", "c" }, row1.ArrayValue);
                Assert.Equal(new int[] { 1, 2, 3 }, row1.MultiMapName);
                Assert.Equal(new string[] { "a", "b" }, row1.MultiMapIndex);

                PrimitiveSheet1 row2 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal(2, row2.IntValue);
                Assert.Equal("b", row2.StringValue);
                Assert.True(row2.BoolValue);
                Assert.Equal(PrimitiveSheet1Enum.Government, row2.EnumValue);
                Assert.Equal(new DateTime(2017, 07, 04), row2.DateValue);
                Assert.Equal(new string[] { "empty" }, row2.ArrayValue);
                Assert.Equal(new int[] { 1, -1, 3 }, row2.MultiMapName);
                Assert.Equal(new string[] { null, null }, row2.MultiMapIndex);

                PrimitiveSheet1 row3 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal(-1, row3.IntValue);
                Assert.Null(row3.StringValue);
                Assert.True(row3.BoolValue);
                Assert.Equal(PrimitiveSheet1Enum.NGO, row3.EnumValue);
                Assert.Equal(new DateTime(10), row3.DateValue);
                Assert.Equal(new string[] { "a" }, row3.ArrayValue);
                Assert.Equal(new int[] { -1, -1, -1 }, row3.MultiMapName);
                Assert.Equal(new string[] { null, "d" }, row3.MultiMapIndex);

                PrimitiveSheet1 row4 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal(-2, row4.IntValue);
                Assert.Equal("d", row4.StringValue);
                Assert.True(row4.BoolValue);
                Assert.Equal(PrimitiveSheet1Enum.Unknown, row4.EnumValue);
                Assert.Equal(new DateTime(20), row4.DateValue);
                Assert.Equal(new string[] { "a", "b", "c", "d", "e" }, row4.ArrayValue);
                Assert.Equal(new int[] { -2, -2, 3 }, row4.MultiMapName);
                Assert.Equal(new string[] { "d", null }, row4.MultiMapIndex);
            }
        }

        public class PrimitiveSheet1
        {
            public int IntValue { get; set; }
            public string StringValue { get; set; }
            public bool BoolValue { get; set; }
            public PrimitiveSheet1Enum EnumValue { get; set; }
            public DateTime DateValue { get; set; }
            public string[] ArrayValue { get; set; }
            public int[] MultiMapName { get; set; }
            public string[] MultiMapIndex { get; set; }
        }

        public enum PrimitiveSheet1Enum
        {
            Unknown = 1,
            Empty,
            Government,
            NGO
        }

        public class PrimitiveSheet1Mapping : ExcelClassMap<PrimitiveSheet1>
        {
            public PrimitiveSheet1Mapping() : base()
            {
                Map(p => p.IntValue)
                    .WithColumnName("Int Value")
                    .WithEmptyFallback(-2)
                    .WithInvalidFallback(-1);

                Map(p => p.StringValue);

                Map(p => p.BoolValue)
                    .WithIndex(2)
                    .WithInvalidFallback(true)
                    .WithEmptyFallback(true);

                Map(p => p.EnumValue)
                    .WithColumnName("Enum Value")
                    .WithMapping(new Dictionary<string, PrimitiveSheet1Enum>
                    {
                        { "Gov't", PrimitiveSheet1Enum.Government }
                    })
                    .WithEmptyFallback(PrimitiveSheet1Enum.Empty)
                    .WithInvalidFallback(PrimitiveSheet1Enum.Unknown);

                Map(p => p.DateValue)
                    .WithAdditionalDateFormats("dd-MM-yyyy")
                    .WithEmptyFallback(new DateTime(10))
                    .WithInvalidFallback(new DateTime(20));

                Map(p => p.ArrayValue)
                    .WithNewDelimiters<DefaultPipeline<string[]>, string[], string>(',', ';')
                    .WithEmptyFallback(new string[] { "empty" });

                MultiMap<int[], int>(p => p.MultiMapName, "SplitValue1", "SplitValue2", "SplitValue3")
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2);

                MultiMap<string[], string>(p => p.MultiMapIndex, 9, 10);
            }
        }
    }
}
