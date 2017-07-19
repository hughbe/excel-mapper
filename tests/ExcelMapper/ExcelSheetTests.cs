using System;
using System.Collections.Generic;
using System.Linq;
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

                Assert.Equal(new string[] { "Int Value", "StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "MappedValue", "TrimmedValue" }, heading.ColumnNames);
            }
        }

        [Fact]
        public void ReadHeading_EmptyColumnName_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("EmptyColumns.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                ExcelHeading heading = sheet.ReadHeading();
                Assert.Same(heading, sheet.Heading);

                Assert.Equal(new string[] { "", "Column2", "", " Column4 " }, heading.ColumnNames);
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
        public void ReadHeading_NoRows_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.ReadSheet();

                ExcelSheet emptySheet = importer.ReadSheet();
                Assert.Throws<ExcelMappingException>(() => emptySheet.ReadHeading());
            }
        }

        [Fact]
        public void ReadRows_NotReadHeading_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();

                IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
                Assert.Equal(new string[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());

                Assert.NotNull(sheet.Heading);
                Assert.True(sheet.HasHeading);
            }
        }

        [Fact]
        public void ReadRows_ReadHeading_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
                Assert.Equal(new string[] { "value", "  value  ", null, "value"  }, rows.Select(p => p.Value).ToArray());

                Assert.NotNull(sheet.Heading);
                Assert.True(sheet.HasHeading);
            }
        }

        private class StringValue
        {
            public string Value { get; set; }
        }

        [Fact]
        public void ReadRow_CantMapType_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Helpers.IListInterface>());
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IConvertible>());
            }
        }

        [Fact]
        public void ReadRow_NoMoreRows_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<PrimitiveSheet1Map>();

                ExcelSheet sheet = importer.ReadSheet();
                Assert.NotEmpty(sheet.ReadRows<PrimitiveSheet1>().ToArray());

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<PrimitiveSheet1>());
                Assert.False(sheet.TryReadRow(out PrimitiveSheet1 row));
                Assert.Null(row);
            }
        }

        [Fact]
        public void ReadRow_Map_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<PrimitiveSheet1Map>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                PrimitiveSheet1 row1 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Null(row1.TrimmedValue);

                PrimitiveSheet1 row2 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal("a", row2.TrimmedValue);

                PrimitiveSheet1 row3 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Null(row3.TrimmedValue);

                PrimitiveSheet1 row4 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal("c", row4.TrimmedValue);
            }
        }

        private class PrimitiveSheet1
        {
            public string TrimmedValue { get; set; }
        }

        private class PrimitiveSheet1Map : ExcelClassMap<PrimitiveSheet1>
        {
            public PrimitiveSheet1Map()
            {

                Map(p => p.TrimmedValue)
                    .WithTrim();
            }
        }

        [Fact]
        public void ReadRow_MultiMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("MultiMap.xlsx"))
            {
                importer.Configuration.RegisterClassMap(new MultiMapRowMap());

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                MultiMapRow row1 = sheet.ReadRow<MultiMapRow>();
                Assert.Equal(new int[] { 1, 2, 3 }, row1.MultiMapName);
                Assert.Equal(new string[] { "a", "b" }, row1.MultiMapIndex);
                Assert.Equal(new int[] { 1, 2 }, row1.IEnumerableInt);
                Assert.Equal(new bool[] { true, false }, row1.ICollectionBool);
                Assert.Equal(new string[] { "a", "b" }, row1.IListString);
                Assert.Equal(new string[] { "1", "2" }, row1.ListString);
                Assert.Equal(new string[] { "1", "2" }, row1._concreteICollection);

                MultiMapRow row2 = sheet.ReadRow<MultiMapRow>();
                Assert.Equal(new int[] { 1, -1, 3 }, row2.MultiMapName);
                Assert.Equal(new string[] { null, null }, row2.MultiMapIndex);
                Assert.Equal(new int[] { 0, 0 }, row2.IEnumerableInt);
                Assert.Equal(new bool[] { false, true }, row2.ICollectionBool);
                Assert.Equal(new string[] { "c", "d" }, row2.IListString);
                Assert.Equal(new string[] { "3", "4" }, row2.ListString);
                Assert.Equal(new string[] { "3", "4" }, row2._concreteICollection);

                MultiMapRow row3 = sheet.ReadRow<MultiMapRow>();
                Assert.Equal(new int[] { -1, -1, -1 }, row3.MultiMapName);
                Assert.Equal(new string[] { null, "d" }, row3.MultiMapIndex);
                Assert.Equal(new int[] { 5, 6 }, row3.IEnumerableInt);
                Assert.Equal(new bool[] { false, false }, row3.ICollectionBool);
                Assert.Equal(new string[] { "e", "f" }, row3.IListString);
                Assert.Equal(new string[] { "5", "6" }, row3.ListString);
                Assert.Equal(new string[] { "5", "6" }, row3._concreteICollection);

                MultiMapRow row4 = sheet.ReadRow<MultiMapRow>();
                Assert.Equal(new int[] { -2, -2, 3 }, row4.MultiMapName);
                Assert.Equal(new string[] { "d", null }, row4.MultiMapIndex);
                Assert.Equal(new int[] { 7, 8 }, row4.IEnumerableInt);
                Assert.Equal(new bool[] { false, true }, row4.ICollectionBool);
                Assert.Equal(new string[] { "g", "h" }, row4.IListString);
                Assert.Equal(new string[] { "7", "8" }, row4.ListString);
                Assert.Equal(new string[] { "7", "8" }, row4._concreteICollection);
            }
        }

        private class MultiMapRow
        {
            public int[] MultiMapName { get; set; }
            public string[] MultiMapIndex { get; set; }
            public IEnumerable<int> IEnumerableInt { get; set; }
            public ICollection<bool> ICollectionBool { get; set; }
            public IList<string> IListString { get; set; }
            public List<string> ListString { get; set; }
#pragma warning disable 0649
            public SortedSet<string> _concreteICollection;
#pragma warning restore 0649
        }

        private class MultiMapRowMap : ExcelClassMap<MultiMapRow>
        {
            public MultiMapRowMap()
            {
                Map(p => p.MultiMapName)
                    .WithColumnNames("MultiMapName1", "MultiMapName2", "MultiMapName3")
                    .WithElementMap(e => e
                        .WithEmptyFallback(-1)
                        .WithInvalidFallback(-2)
                    );

                Map(p => p.MultiMapIndex)
                    .WithColumnIndices(3, 4);

                Map(p => p.IEnumerableInt)
                    .WithColumnNames(new List<string> { "IEnumerableInt1", "IEnumerableInt2" })
                    .WithElementMap(e => e
                        .WithValueFallback(default(int))
                    );

                Map(p => p.ICollectionBool)
                    .WithColumnIndices(new List<int> { 7, 8 })
                    .WithElementMap(e => e
                        .WithValueFallback(default(bool))
                    );

                Map(p => p.IListString)
                    .WithColumnNames("IListString1", "IListString2");

                Map(p => p.ListString)
                    .WithColumnNames("ListString1", "ListString2");

                Map<string>(p => p._concreteICollection)
                    .WithColumnNames("ListString1", "ListString2");
            }
        }

        [Fact]
        public void ReadRow_NullableValues_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("EmptyValues.xlsx"))
            {
                importer.Configuration.RegisterClassMap(new NullableValuesMap());

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                NullableValues row1 = sheet.ReadRow<NullableValues>();
                Assert.Equal(new int?[] { null, null }, row1.ArrayValue);

                NullableValues row2 = sheet.ReadRow<NullableValues>();
                Assert.Equal(new int?[] { 1, 2 }, row2.ArrayValue);
            }
        }

        private class NullableValues
        {
            public int?[] ArrayValue { get; set; }
        }

        private class NullableValuesMap : ExcelClassMap<NullableValues>
        {
            public NullableValuesMap() : base(FallbackStrategy.SetToDefaultValue)
            {
                Map(n => n.ArrayValue)
                    .WithColumnNames("ArrayValue1", "ArrayValue2");
            }
        }
    }
}
