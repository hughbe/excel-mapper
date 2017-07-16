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
        public void ReadRows_NotReadHeading_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterMapping<PrimitiveSheet1Mapping>();

                ExcelSheet sheet = importer.ReadSheet();
                IEnumerable<PrimitiveSheet1> rows = sheet.ReadRows<PrimitiveSheet1>().ToArray();
                Assert.Equal(new int[] { 1, 2, -1, -2 }, rows.Select(p => p.IntValue));

                Assert.NotNull(sheet.Heading);
                Assert.True(sheet.HasHeading);
            }
        }

        [Fact]
        public void ReadRows_ReadHeading_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterMapping<PrimitiveSheet1Mapping>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                IEnumerable<PrimitiveSheet1> rows = sheet.ReadRows<PrimitiveSheet1>().ToArray();
                Assert.Equal(new int[] { 1, 2, -1, -2 }, rows.Select(p => p.IntValue));

                Assert.NotNull(sheet.Heading);
                Assert.True(sheet.HasHeading);
            }
        }

        [Fact]
        public void ReadRow_NoMoreRows_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterMapping<PrimitiveSheet1Mapping>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadRows<PrimitiveSheet1>().ToArray();

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
                importer.Configuration.RegisterMapping<PrimitiveSheet1Mapping>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                PrimitiveSheet1 row1 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal(1, row1.IntValue);
                Assert.Equal("a", row1.StringValue);
                Assert.True(row1.BoolValue);
                Assert.Equal(PrimitiveSheet1Enum.Government, row1.EnumValue);
                Assert.Equal(new DateTime(2017, 07, 04), row1.DateValue);
                Assert.Equal(new DateTime(2017, 07, 04), row1.DateValue2);
                Assert.Equal(new string[] { "a", "b", "c" }, row1.ArrayValue);
                Assert.Equal("MappedA", row1.MappedValue);
                Assert.Equal(string.Empty, row1.TrimmedValue);

                PrimitiveSheet1 row2 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal(2, row2.IntValue);
                Assert.Equal("b", row2.StringValue);
                Assert.True(row2.BoolValue);
                Assert.Equal(PrimitiveSheet1Enum.Government, row2.EnumValue);
                Assert.Equal(new DateTime(2017, 07, 04), row2.DateValue);
                Assert.Equal(new DateTime(2017, 07, 04), row2.DateValue2);
                Assert.Equal(new string[] { }, row2.ArrayValue);
                Assert.Equal("MappedB", row2.MappedValue);
                Assert.Equal("a", row2.TrimmedValue);

                PrimitiveSheet1 row3 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal(-1, row3.IntValue);
                Assert.Null(row3.StringValue);
                Assert.True(row3.BoolValue);
                Assert.Equal(PrimitiveSheet1Enum.NGO, row3.EnumValue);
                Assert.Equal(new DateTime(10), row3.DateValue);
                Assert.Equal(new DateTime(10), row3.DateValue2);
                Assert.Equal(new string[] { "a" }, row3.ArrayValue);
                Assert.Equal("MappedB", row3.MappedValue);
                Assert.Null(row3.TrimmedValue);

                PrimitiveSheet1 row4 = sheet.ReadRow<PrimitiveSheet1>();
                Assert.Equal(-2, row4.IntValue);
                Assert.Equal("d", row4.StringValue);
                Assert.True(row4.BoolValue);
                Assert.Equal(PrimitiveSheet1Enum.Unknown, row4.EnumValue);
                Assert.Equal(new DateTime(20), row4.DateValue);
                Assert.Equal(new DateTime(20), row4.DateValue2);
                Assert.Equal(new string[] { "a", "b", "c", "d", "e" }, row4.ArrayValue);
                Assert.Equal("D", row4.MappedValue);
                Assert.Equal("c", row4.TrimmedValue);
            }
        }

        public class PrimitiveSheet1
        {
            public int IntValue { get; set; }
            public string StringValue { get; set; }
            public bool BoolValue { get; set; }
            public PrimitiveSheet1Enum EnumValue { get; set; }
            public DateTime DateValue { get; set; }
            public DateTime DateValue2 { get; set; }
            public string[] ArrayValue { get; set; }
            public string MappedValue { get; set; }
            public string TrimmedValue { get; set; }
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
                    .WithColumnIndex(2)
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
                    .WithDateFormats("G", "dd-MM-yyyy")
                    .WithEmptyFallback(new DateTime(10))
                    .WithInvalidFallback(new DateTime(20));

                Map(p => p.DateValue2)
                    .WithColumnName("DateValue")
                    .WithDateFormats(new List<string> { "G", "dd-MM-yyyy" })
                    .WithEmptyFallback(new DateTime(10))
                    .WithInvalidFallback(new DateTime(20));

                Map(p => p.ArrayValue)
                    .WithSeparators(',', ';')
                    .WithElementMapping(e => e
                        .WithEmptyFallback("empty")
                    );

                Map(p => p.MappedValue)
                    .WithMapping(new Dictionary<string, string>
                    {
                        { "a", "MappedA" },
                        { "b", "MappedB" }
                    }, StringComparer.OrdinalIgnoreCase);

                Map(p => p.TrimmedValue)
                    .WithTrim();
            }
        }

        [Fact]
        public void ReadRow_MultiMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("MultiMap.xlsx"))
            {
                importer.Configuration.RegisterMapping(new MultiMapRowMapping());

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

        public class MultiMapRow
        {
            public int[] MultiMapName { get; set; }
            public string[] MultiMapIndex { get; set; }
            public IEnumerable<int> IEnumerableInt { get; set; }
            public ICollection<bool> ICollectionBool { get; set; }
            public IList<string> IListString { get; set; }
            public List<string> ListString { get; set; }
            public SortedSet<string> _concreteICollection;
        }

        public class MultiMapRowMapping : ExcelClassMap<MultiMapRow>
        {
            public MultiMapRowMapping() : base()
            {
                Map(p => p.MultiMapName)
                    .WithColumnNames("MultiMapName1", "MultiMapName2", "MultiMapName3")
                    .WithElementMapping((SinglePropertyMapping<int> e) => e
                        .WithEmptyFallback(-1)
                        .WithInvalidFallback(-2)
                    );

                Map(p => p.MultiMapIndex)
                    .WithColumnIndices(3, 4);

                Map(p => p.IEnumerableInt)
                    .WithColumnNames(new List<string> { "IEnumerableInt1", "IEnumerableInt2" })
                    .WithElementMapping(e => e
                        .WithValueFallback(default(int))
                    );

                Map(p => p.ICollectionBool)
                    .WithColumnIndices(new List<int> { 7, 8 })
                    .WithElementMapping((SinglePropertyMapping<bool> e) => e
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
        public void ReadRow_SplitWithSeparator_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("MultiMap.xlsx"))
            {
                importer.Configuration.RegisterMapping(new SplitWithSeparatorMapping());

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                SplitWithSeparatorClass row1 = sheet.ReadRow<SplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3" }, row1.CommaSeparator);

                Assert.Equal(new string[] { "1", "2", "3" }, row1.CommaSeparatorWithColumnName);
                Assert.Equal(new string[] { "1", "2", "3" }, row1.CommaSeparatorWithColumnIndex);

                Assert.Equal(new string[] { "1", "2", "3" }, row1.CommaSeparatorWithColumnNameAcrossMultiColumnNames);
                Assert.Equal(new string[] { "1", "2", "3" }, row1.CommaSeparatorWithColumnNameAcrossMultiColumnIndices);

                Assert.Equal(new string[] { "1", "2", "3" }, row1.CommaSeparatorWithColumnIndexAcrossMultiColumnNames);
                Assert.Equal(new string[] { "1", "2", "3" }, row1.CommaSeparatorWithColumnIndexAcrossMultiColumnIndices);
            }
        }

        public class SplitWithSeparatorClass
        {
            public string[] CommaSeparator { get; set; }
            public string[] CommaSeparatorWithColumnName { get; set; }
            public string[] CommaSeparatorWithColumnIndex { get; set; }

            public string[] CommaSeparatorWithColumnNameAcrossMultiColumnNames { get; set; }
            public string[] CommaSeparatorWithColumnNameAcrossMultiColumnIndices { get; set; }

            public string[] CommaSeparatorWithColumnIndexAcrossMultiColumnNames { get; set; }
            public string[] CommaSeparatorWithColumnIndexAcrossMultiColumnIndices { get; set; }
        }

        public class SplitWithSeparatorMapping : ExcelClassMap<SplitWithSeparatorClass>
        {
            public SplitWithSeparatorMapping() : base()
            {
                Map(p => p.CommaSeparator);

                Map(p => p.CommaSeparatorWithColumnName)
                    .WithColumnName("CommaSeparator");

                Map(p => p.CommaSeparatorWithColumnIndex)
                    .WithColumnIndex(13);

                Map(p => p.CommaSeparatorWithColumnNameAcrossMultiColumnNames)
                    .WithColumnNames("IListString1", "IListString2")
                    .WithColumnName("CommaSeparator");

                Map(p => p.CommaSeparatorWithColumnNameAcrossMultiColumnIndices)
                    .WithColumnIndices(9, 10)
                    .WithColumnName("CommaSeparator");

                Map(p => p.CommaSeparatorWithColumnIndexAcrossMultiColumnNames)
                    .WithColumnNames("IListString1", "IListString2")
                    .WithColumnIndex(13);

                Map(p => p.CommaSeparatorWithColumnIndexAcrossMultiColumnIndices)
                    .WithColumnIndices(9, 10)
                    .WithColumnIndex(13);
            }
        }

        [Fact]
        public void ReadRow_EmptyValueStrategy_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("EmptyValues.xlsx"))
            {
                importer.Configuration.RegisterMapping(new EmptyValueStrategyMapping());

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                EmptyValues row1 = sheet.ReadRow<EmptyValues>();
                Assert.Equal(0, row1.IntValue);
                Assert.Null(row1.StringValue);
                Assert.False(row1.BoolValue);
                Assert.Equal((EmptyValuesEnum)0, row1.EnumValue);
                Assert.Equal(DateTime.MinValue, row1.DateValue);
                Assert.Equal(new int[] { 0, 0 }, row1.ArrayValue);
            }
        }

        public class EmptyValues
        {
            public int IntValue { get; set; }
            public string StringValue { get; set; }
            public bool BoolValue { get; set; }
            public EmptyValuesEnum EnumValue { get; set; }
            public DateTime DateValue { get; set; }
            public int[] ArrayValue { get; set; }
        }

        public enum EmptyValuesEnum
        {
            Test = 1
        }

        public class EmptyValueStrategyMapping : ExcelClassMap<EmptyValues>
        {
            public EmptyValueStrategyMapping() : base(EmptyValueStrategy.SetToDefaultValue)
            {
                Map(e => e.IntValue);
                Map(e => e.StringValue);
                Map(e => e.BoolValue);
                Map(e => e.EnumValue);
                Map(e => e.DateValue);
                Map<int>(e => e.ArrayValue)
                    .WithColumnNames("ArrayValue1", "ArrayValue2");
            }
        }

        [Fact]
        public void ReadRow_NullableValues_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("EmptyValues.xlsx"))
            {
                importer.Configuration.RegisterMapping(new NullableValuesMapping());

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                NullableValues row1 = sheet.ReadRow<NullableValues>();
                Assert.Null(row1.IntValue);
                Assert.Null(row1.BoolValue);
                Assert.Null(row1.EnumValue);
                Assert.Null(row1.DateValue);
                Assert.Null(row1.DateValueWithFormats);
                Assert.Equal(new int?[] { null, null }, row1.ArrayValue);

                NullableValues row2 = sheet.ReadRow<NullableValues>();
                Assert.Equal(1, row2.IntValue);
                Assert.True(row2.BoolValue);
                Assert.Equal(NullableValuesEnum.Test, row2.EnumValue);
                Assert.Equal(new DateTime(2017, 07, 05), row2.DateValue);
                Assert.Equal(new DateTime(2017, 07, 05), row2.DateValueWithFormats);
                Assert.Equal(new int?[] { 1, 2 }, row2.ArrayValue);
            }
        }

        public class NullableValues
        {
            public int? IntValue { get; set; }
            public bool? BoolValue { get; set; }
            public NullableValuesEnum? EnumValue { get; set; }
            public DateTime? DateValue { get; set; }
            public DateTime? DateValueWithFormats { get; set; }
            public int?[] ArrayValue { get; set; }
        }

        public enum NullableValuesEnum
        {
            Test = 1
        }

        public class NullableValuesMapping : ExcelClassMap<NullableValues>
        {
            public NullableValuesMapping() : base(EmptyValueStrategy.SetToDefaultValue)
            {
                Map(n => n.IntValue);
                Map(n => n.BoolValue);
                Map(n => n.EnumValue);
                Map(n => n.DateValue);
                Map(n => n.DateValueWithFormats)
                    .WithColumnName("DateValue")
                    .WithDateFormats("G");
                Map(n => n.ArrayValue)
                    .WithColumnNames("ArrayValue1", "ArrayValue2");
            }
        }

        [Fact]
        public void ReadRow_OptionalMapping_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterMapping<OptionalValueMapping>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                OptionalValue row1 = sheet.ReadRow<OptionalValue>();
                Assert.Equal(-1, row1.NoSuchColumnNoName);
                Assert.Equal(-2, row1.NoSuchColumnWithNameBefore);
                Assert.Equal(-3, row1.NoSuchColumnWithNameAfter);
                Assert.Equal(-4, row1.NoSuchColumnWithIndexBefore);
                Assert.Equal(-5, row1.NoSuchColumnWithIndexAfter);
            }
        }

        public class OptionalValue
        {
            public int NoSuchColumnNoName { get; set; }

            public int NoSuchColumnWithNameBefore { get; set; }
            public int NoSuchColumnWithNameAfter { get; set; }

            public int NoSuchColumnWithIndexBefore { get; set; }
            public int NoSuchColumnWithIndexAfter { get; set; }
        }

        public class OptionalValueMapping : ExcelClassMap<OptionalValue>
        {
            public OptionalValueMapping()
            {
                Map(v => v.NoSuchColumnNoName)
                    .MakeOptional()
                    .WithEmptyFallback(-1);

                Map(v => v.NoSuchColumnWithNameBefore)
                    .WithColumnName("NoSuchColumn")
                    .MakeOptional()
                    .WithEmptyFallback(-2);

                Map(v => v.NoSuchColumnWithNameAfter)
                    .MakeOptional()
                    .WithColumnName("NoSuchColumn")
                    .WithEmptyFallback(-3);

                Map(v => v.NoSuchColumnWithIndexBefore)
                    .WithColumnIndex(10)
                    .MakeOptional()
                    .WithEmptyFallback(-4);

                Map(v => v.NoSuchColumnWithIndexAfter)
                    .MakeOptional()
                    .WithColumnIndex(10)
                    .WithEmptyFallback(-5);
            }
        }

        [Fact]
        public void ReadRow_ConvertUsingMapping_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterMapping<ConvertUsingValueMapping>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                ConvertUsingValue row1 = sheet.ReadRow<ConvertUsingValue>();
                Assert.Equal("aextra", row1.StringValue);
            }
        }

        public class ConvertUsingValue
        {
            public string StringValue { get; set; }
        }

        public class ConvertUsingValueMapping : ExcelClassMap<ConvertUsingValue>
        {
            public ConvertUsingValueMapping()
            {
                Map(c => c.StringValue)
                    .WithConverter(s => s + "extra");
            }
        }
    }
}
