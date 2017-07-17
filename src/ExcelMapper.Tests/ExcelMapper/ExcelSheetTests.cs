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
                importer.Configuration.RegisterClassMap<PrimitiveSheet1Map>();

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
                importer.Configuration.RegisterClassMap<PrimitiveSheet1Map>();

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
                importer.Configuration.RegisterClassMap<PrimitiveSheet1Map>();

                ExcelSheet sheet = importer.ReadSheet();
                Assert.NotEmpty(sheet.ReadRows<PrimitiveSheet1>().ToArray());

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<PrimitiveSheet1>());
                Assert.False(sheet.TryReadRow(out PrimitiveSheet1 row));
                Assert.Null(row);
            }
        }

        [Fact]
        public void ReadRow_ObjectValueWithDefaultMap_Success()
        {
            using (var importer = Helpers.GetImporter("Objects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ObjectValueDefaultMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                ObjectValue row1 = sheet.ReadRow<ObjectValue>();
                Assert.Equal("value", row1.Value);

                ObjectValue row2 = sheet.ReadRow<ObjectValue>();
                Assert.Null(row2.Value);
            }
        }

        public class ObjectValue
        {
            public object Value { get; set; }
        }

        private class ObjectValueDefaultMap : ExcelClassMap<ObjectValue>
        {
            public ObjectValueDefaultMap()
            {
                Map(o => o.Value);
            }
        }

        [Fact]
        public void ReadRow_ObjectValueWithFallbackMap_Success()
        {
            using (var importer = Helpers.GetImporter("Objects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ObjectValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                ObjectValue row1 = sheet.ReadRow<ObjectValue>();
                Assert.Equal("value", row1.Value);

                ObjectValue row2 = sheet.ReadRow<ObjectValue>();
                Assert.Equal("empty", row2.Value);
            }
        }

        private class ObjectValueFallbackMap : ExcelClassMap<ObjectValue>
        {
            public ObjectValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback("empty")
                    .WithInvalidFallback("invalid");
            }
        }

        [Fact]
        public void ReadRow_StringValueWithDefaultMap_Success()
        {
            using (var importer = Helpers.GetImporter("Objects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<StringValueDefaultMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                StringValue row1 = sheet.ReadRow<StringValue>();
                Assert.Equal("value", row1.Value);

                StringValue row2 = sheet.ReadRow<StringValue>();
                Assert.Null(row2.Value);
            }
        }

        public class StringValue
        {
            public string Value { get; set; }
        }

        private class StringValueDefaultMap : ExcelClassMap<StringValue>
        {
            public StringValueDefaultMap()
            {
                Map(o => o.Value);
            }
        }

        [Fact]
        public void ReadRow_StringValueWithFallbackMap_Success()
        {
            using (var importer = Helpers.GetImporter("Objects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<StringValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                StringValue row1 = sheet.ReadRow<StringValue>();
                Assert.Equal("value", row1.Value);

                StringValue row2 = sheet.ReadRow<StringValue>();
                Assert.Equal("empty", row2.Value);
            }
        }

        private class StringValueFallbackMap : ExcelClassMap<StringValue>
        {
            public StringValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback("empty")
                    .WithInvalidFallback("invalid");
            }
        }

        [Fact]
        public void ReadRow_IConvertibleValueWithDefaultMap_Success()
        {
            using (var importer = Helpers.GetImporter("Objects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ConvertibleValueDefaultMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                ConvertibleValue row1 = sheet.ReadRow<ConvertibleValue>();
                Assert.Equal("value", row1.Value);

                ConvertibleValue row2 = sheet.ReadRow<ConvertibleValue>();
                Assert.Null(row2.Value);
            }
        }

        private class ConvertibleValue
        {
            public IConvertible Value { get; set; }
        }

        private class ConvertibleValueDefaultMap : ExcelClassMap<ConvertibleValue>
        {
            public ConvertibleValueDefaultMap()
            {
                Map(o => o.Value);
            }
        }

        [Fact]
        public void ReadRow_IConvertibleValueWithFallbackMap_Success()
        {
            using (var importer = Helpers.GetImporter("Objects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ConvertibleValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                ConvertibleValue row1 = sheet.ReadRow<ConvertibleValue>();
                Assert.Equal("value", row1.Value);

                ConvertibleValue row2 = sheet.ReadRow<ConvertibleValue>();
                Assert.Equal("empty", row2.Value);
            }
        }

        private class ConvertibleValueFallbackMap : ExcelClassMap<ConvertibleValue>
        {
            public ConvertibleValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback("empty")
                    .WithInvalidFallback("invalid");
            }
        }

        [Fact]
        public void ReadRow_DoubleValueWithDefaultMap_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<DoubleValueDefaultMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                DoubleValue row1 = sheet.ReadRow<DoubleValue>();
                Assert.Equal(2.2345, row1.Value);

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleValue>());
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleValue>());
            }
        }

        private class DoubleValue
        {
            public double Value { get; set; }
        }

        private class DoubleValueDefaultMap : ExcelClassMap<DoubleValue>
        {
            public DoubleValueDefaultMap()
            {
                Map(o => o.Value);
            }
        }

        [Fact]
        public void ReadRow_DoubleValueWithFallbackMap_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<DoubleValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                DoubleValue row1 = sheet.ReadRow<DoubleValue>();
                Assert.Equal(2.2345, row1.Value);

                DoubleValue row2 = sheet.ReadRow<DoubleValue>();
                Assert.Equal(-10, row2.Value);

                DoubleValue row3 = sheet.ReadRow<DoubleValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        private class DoubleValueFallbackMap : ExcelClassMap<DoubleValue>
        {
            public DoubleValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10)
                    .WithInvalidFallback(10);
            }
        }

        [Fact]
        public void ReadRow_FloatValueWithDefaultMap_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<FloatValueDefaultMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                FloatValue row1 = sheet.ReadRow<FloatValue>();
                Assert.Equal(2.2345f, row1.Value);

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatValue>());
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatValue>());
            }
        }

        private class FloatValue
        {
            public float Value { get; set; }
        }

        private class FloatValueDefaultMap : ExcelClassMap<FloatValue>
        {
            public FloatValueDefaultMap()
            {
                Map(o => o.Value);
            }
        }

        [Fact]
        public void ReadRow_FloatValueWithFallbackMap_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<FloatValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                FloatValue row1 = sheet.ReadRow<FloatValue>();
                Assert.Equal(2.2345f, row1.Value);

                FloatValue row2 = sheet.ReadRow<FloatValue>();
                Assert.Equal(-10, row2.Value);

                FloatValue row3 = sheet.ReadRow<FloatValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        private class FloatValueFallbackMap : ExcelClassMap<FloatValue>
        {
            public FloatValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10)
                    .WithInvalidFallback(10);
            }
        }

        [Fact]
        public void ReadRow_DecimalValueWithDefaultMap_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<DecimalValueDefaultMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                DecimalValue row1 = sheet.ReadRow<DecimalValue>();
                Assert.Equal(2.2345m, row1.Value);

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());
            }
        }

        private class DecimalValue
        {
            public decimal Value { get; set; }
        }

        private class DecimalValueDefaultMap : ExcelClassMap<DecimalValue>
        {
            public DecimalValueDefaultMap()
            {
                Map(o => o.Value);
            }
        }

        [Fact]
        public void ReadRow_DecimalValueWithFallbackMap_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<DecimalValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                DecimalValue row1 = sheet.ReadRow<DecimalValue>();
                Assert.Equal(2.2345m, row1.Value);

                DecimalValue row2 = sheet.ReadRow<DecimalValue>();
                Assert.Equal(-10, row2.Value);

                DecimalValue row3 = sheet.ReadRow<DecimalValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        private class DecimalValueFallbackMap : ExcelClassMap<DecimalValue>
        {
            public DecimalValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10m)
                    .WithInvalidFallback(10m);
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

        private class PrimitiveSheet1
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

        private enum PrimitiveSheet1Enum
        {
            Unknown = 1,
            Empty,
            Government,
            NGO
        }

        private class PrimitiveSheet1Map : ExcelClassMap<PrimitiveSheet1>
        {
            public PrimitiveSheet1Map()
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
                    .WithElementMap(e => e
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
            public SortedSet<string> _concreteICollection;
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
        public void ReadRow_SplitWithSeparator_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("MultiMap.xlsx"))
            {
                importer.Configuration.RegisterClassMap(new SplitWithSeparatorMap());

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

        public class SplitWithSeparatorMap : ExcelClassMap<SplitWithSeparatorClass>
        {
            public SplitWithSeparatorMap()
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
                importer.Configuration.RegisterClassMap(new EmptyValueStrategyMap());

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

        public class EmptyValueStrategyMap : ExcelClassMap<EmptyValues>
        {
            public EmptyValueStrategyMap() : base(FallbackStrategy.SetToDefaultValue)
            {
                Map(e => e.IntValue);
                Map(e => e.StringValue);
                Map(e => e.BoolValue);
                Map(e => e.EnumValue);
                Map(e => e.DateValue);
                Map(e => e.ArrayValue)
                    .WithColumnNames("ArrayValue1", "ArrayValue2");
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

        private class NullableValues
        {
            public int? IntValue { get; set; }
            public bool? BoolValue { get; set; }
            public NullableValuesEnum? EnumValue { get; set; }
            public DateTime? DateValue { get; set; }
            public DateTime? DateValueWithFormats { get; set; }
            public int?[] ArrayValue { get; set; }
        }

        private enum NullableValuesEnum
        {
            Test = 1
        }

        private class NullableValuesMap : ExcelClassMap<NullableValues>
        {
            public NullableValuesMap() : base(FallbackStrategy.SetToDefaultValue)
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
        public void ReadRow_OptionalMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<OptionalValueMap>();

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

        private class OptionalValue
        {
            public int NoSuchColumnNoName { get; set; }

            public int NoSuchColumnWithNameBefore { get; set; }
            public int NoSuchColumnWithNameAfter { get; set; }

            public int NoSuchColumnWithIndexBefore { get; set; }
            public int NoSuchColumnWithIndexAfter { get; set; }
        }

        private class OptionalValueMap : ExcelClassMap<OptionalValue>
        {
            public OptionalValueMap()
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
        public void ReadRow_ConvertUsingMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ConvertUsingValueMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                ConvertUsingValue row1 = sheet.ReadRow<ConvertUsingValue>();
                Assert.Equal("aextra", row1.StringValue);
            }
        }

        private class ConvertUsingValue
        {
            public string StringValue { get; set; }
        }

        private class ConvertUsingValueMap : ExcelClassMap<ConvertUsingValue>
        {
            public ConvertUsingValueMap()
            {
                Map(c => c.StringValue)
                    .WithConverter(s => s + "extra");
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

        [Fact]
        public void ReadRow_UriWithDefaultFallback_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Uris.xlsx"))
            {
                importer.Configuration.RegisterClassMap<UriDefaultFallbackMap>();

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

        private class UriDefaultFallbackMap : ExcelClassMap<UriValue>
        {
            public UriDefaultFallbackMap()
            {
                Map(u => u.Uri);
            }
        }

        [Fact]
        public void ReadRow_Object_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("NestedObjects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ObjectValueDefaultClassMapMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
                Assert.Equal("a", row1.SubValue1.StringValue);
                Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
                Assert.Equal(1, row1.SubValue2.IntValue);
                Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
                Assert.Equal("c", row1.SubValue2.SubValue.SubString);
            }
        }

        private class NestedObjectValue
        {
            public SubValue1 SubValue1 { get; set; }
            public SubValue2 SubValue2 { get; set; }
        }

        private class SubValue1
        {
            public string StringValue { get; set; }
            public string[] SplitStringValue { get; set; }
        }

        private class SubValue2
        {
            public int IntValue { get; set; }
            public SubValue3 SubValue { get; set; }
        }

        private class SubValue3
        {
            public string SubString { get; set; }
            public int SubInt { get; set; }
        }

        private class ObjectValueDefaultClassMapMap : ExcelClassMap<NestedObjectValue>
        {
            public ObjectValueDefaultClassMapMap()
            {
                MapObject(p => p.SubValue1);
                MapObject(p => p.SubValue2);
            }
        }

        [Fact]
        public void ReadRow_ObjectWithCustomClassMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("NestedObjects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ObjectValueCustomClassMapMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
                Assert.Equal("a", row1.SubValue1.StringValue);
                Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
                Assert.Equal(1, row1.SubValue2.IntValue);
                Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
                Assert.Equal("c", row1.SubValue2.SubValue.SubString);
            }
        }

        private class ObjectValueCustomClassMapMap : ExcelClassMap<NestedObjectValue>
        {
            public ObjectValueCustomClassMapMap()
            {
                MapObject(p => p.SubValue1).WithClassMap(m =>
                {
                    m.Map(s => s.StringValue);
                    m.Map(s => s.SplitStringValue);
                });

                MapObject(p => p.SubValue2).WithClassMap(new SubValueMap());
            }
        }

        private class SubValueMap : ExcelClassMap<SubValue2>
        {
            public SubValueMap()
            {
                Map(s => s.IntValue);

                MapObject(s => s.SubValue);
            }
        }

        [Fact]
        public void ReadRow_ObjectInnerMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("NestedObjects.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ObjectValueInnerMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
                Assert.Equal("a", row1.SubValue1.StringValue);
                Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
                Assert.Equal(1, row1.SubValue2.IntValue);
                Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
                Assert.Equal("c", row1.SubValue2.SubValue.SubString);
            }
        }

        private class ObjectValueInnerMap : ExcelClassMap<NestedObjectValue>
        {
            public ObjectValueInnerMap()
            {
                Map(p => p.SubValue1.StringValue);
                Map(p => p.SubValue1.SplitStringValue);
                Map(p => p.SubValue2.IntValue);
                Map(p => p.SubValue2.SubValue.SubInt);
                Map(p => p.SubValue2.SubValue.SubString);
            }
        }
    }
}
