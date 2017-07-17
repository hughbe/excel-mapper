using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Mappers;
using ExcelMapper.Mappings.Readers;
using ExcelMapper.Mappings.Transformers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class SinglePropertyMappingExtensionsTests : ExcelClassMap<Helpers.TestClass>
    {
        [Fact]
        public void WithColumnName_ValidColumnName_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            ColumnNameValueReader reader = Assert.IsType<ColumnNameValueReader>(propertyMap.CellReader);
            Assert.Equal("ColumnName", reader.ColumnName);
        }

        [Fact]
        public void WithColumnName_OptionalColumn_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            OptionalCellValueReader reader = Assert.IsType<OptionalCellValueReader>(propertyMap.CellReader);
            ColumnNameValueReader innerReader = Assert.IsType<ColumnNameValueReader>(reader.InnerReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("columnName", () => propertyMap.WithColumnName(null));
        }

        [Fact]
        public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentException>("columnName", () => propertyMap.WithColumnName(string.Empty));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_ValidColumnIndex_Success(int columnIndex)
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

            ColumnIndexValueReader reader = Assert.IsType<ColumnIndexValueReader>(propertyMap.CellReader);
            Assert.Equal(columnIndex, reader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_OptionalColumn_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(1));

            OptionalCellValueReader reader = Assert.IsType<OptionalCellValueReader>(propertyMap.CellReader);
            ColumnIndexValueReader innerReader = Assert.IsType<ColumnIndexValueReader>(reader.InnerReader);
            Assert.Equal(1, innerReader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => propertyMap.WithColumnIndex(-1));
        }

        [Fact]
        public void WithMapper_ValidMapper_Success()
        {
            var reader = new ColumnNameValueReader("ColumnName");
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithReader(reader));

            Assert.Same(reader, propertyMap.CellReader);
        }

        [Fact]
        public void WithReader_OptionalColumn_Success()
        {
            var innerReader = new ColumnNameValueReader("ColumnName");

            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.Same(propertyMap, propertyMap.WithReader(innerReader));

            OptionalCellValueReader reader = Assert.IsType<OptionalCellValueReader>(propertyMap.CellReader);
            Assert.Same(innerReader, reader.InnerReader);
        }

        [Fact]
        public void WithReader_NullReader_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("reader", () => propertyMap.WithReader(null));
        }

        [Fact]
        public void WithMappingItems_ValidMappingItems_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            ICellValueMapper mapper1 = Assert.Single(propertyMap.CellValueMappers);
            ICellValueMapper mapper2 = new BoolMapper(); ;

            Assert.Same(propertyMap, propertyMap.WithMappingItems(mapper2));
            Assert.Equal(new ICellValueMapper[] { mapper1, mapper2 }, propertyMap.CellValueMappers);
        }

        [Fact]
        public void WithMappingItems_NullMappingItems_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("mappings", () => propertyMap.WithMappingItems(null));
        }

        [Fact]
        public void WithMapping_ValidReader_Success()
        {
            var dictionaryMapping = new Dictionary<string, DateTime> { { "key", DateTime.MinValue } };
            StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;

            SingleExcelPropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithMapping(dictionaryMapping, comparer));

            DictionaryMapper<DateTime> item = propertyMap.CellValueMappers.OfType<DictionaryMapper<DateTime>>().Single();
            Assert.NotSame(dictionaryMapping, item.MappingDictionary);
            Assert.Equal(dictionaryMapping, item.MappingDictionary);

            Assert.Same(comparer, Assert.IsType<Dictionary<string, DateTime>>(item.MappingDictionary).Comparer);
        }

        [Fact]
        public void WithMapping_NullMapping_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("mappingDictionary", () => propertyMap.WithMapping((Dictionary<string, string>)null));
        }

        [Fact]
        public void MakeOptional_HasMapper_ReturnsExpected()
        {
            var innerReader = new ColumnIndexValueReader(1);
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value).WithReader(innerReader);
            Assert.Same(propertyMap, propertyMap.MakeOptional());

            OptionalCellValueReader reader = Assert.IsType<OptionalCellValueReader>(propertyMap.CellReader);
            Assert.Same(innerReader, reader.InnerReader);
        }

        [Fact]
        public void MakeOptional_AlreadyOptional_ThrowsExcelMappingException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value).WithColumnIndex(1).MakeOptional();

            Assert.Throws<ExcelMappingException>(() => propertyMap.MakeOptional());
        }

        [Fact]
        public void WithTrim_Invoke_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithTrim());

            ICellValueTransformer transformer = Assert.Single(propertyMap.CellValueTransformers);
            Assert.IsType<TrimCellValueTransformer>(transformer);
        }

        public static IEnumerable<object[]> Formats_TestData()
        {
            yield return new object[] { new string[] { "1" } };
            yield return new object[] { new string[] { "g", "yyyy-MM-dd" } };
            yield return new object[] { new List<string> { "g", "yyyy-MM-dd" } };
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_AutoMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            SingleExcelPropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NotMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            SingleExcelPropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            propertyMap.RemoveCellValueMapper(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_AutoMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            SingleExcelPropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NotMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            SingleExcelPropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            propertyMap.RemoveCellValueMapper(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Fact]
        public void WithDateFormats_NullFormats_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<DateTime> propertyMap = Map(t => t.DateValue);

            Assert.Throws<ArgumentNullException>("formats", () => propertyMap.WithDateFormats(null));
            Assert.Throws<ArgumentNullException>("formats", () => propertyMap.WithDateFormats((IEnumerable<string>)null));
        }

        [Fact]
        public void WithDateFormats_EmptyFormats_ThrowsArgumentException()
        {
            SingleExcelPropertyMap<DateTime> propertyMap = Map(t => t.DateValue);

            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new string[0]));
            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new List<string>()));
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableAutoMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            SingleExcelPropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableNotMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            SingleExcelPropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            propertyMap.RemoveCellValueMapper(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableAutoMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            SingleExcelPropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableNotMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            SingleExcelPropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            propertyMap.RemoveCellValueMapper(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Fact]
        public void WithDateFormats_NullableNullFormats_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<DateTime?> mapping = Map(t => t.NullableDateValue);

            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats(null));
            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats((IEnumerable<string>)null));
        }

        [Fact]
        public void WithDateFormats_NullableEmptyFormats_ThrowsArgumentException()
        {
            SingleExcelPropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);

            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new string[0]));
            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new List<string>()));
        }

        [Fact]
        public void WithConverter_SuccessConverter_ReturnsExpected()
        {
            ConvertUsingSimpleMappingDelegate<string> converter = stringValue =>
            {
                Assert.Equal("stringValue", stringValue);
                return "abc";
            };

            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithConverter(converter));
            ConvertUsingMapper item = propertyMap.CellValueMappers.OfType<ConvertUsingMapper>().Single();

            object value = null;
            PropertyMappingResultType result = item.Converter(new ReadCellValueResult(-1, "stringValue"), ref value);
            Assert.Equal(PropertyMappingResultType.Success, result);
            Assert.Equal("abc", value);
        }

        [Fact]
        public void WithConverter_InvalidConverter_ReturnsExpected()
        {
            ConvertUsingSimpleMappingDelegate<string> converter = stringValue =>
            {
                Assert.Equal("stringValue", stringValue);
                throw new NotSupportedException();
            };

            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithConverter(converter));
            ConvertUsingMapper item = propertyMap.CellValueMappers.OfType<ConvertUsingMapper>().Single();

            object value = 1;
            PropertyMappingResultType result = item.Converter(new ReadCellValueResult(-1, "stringValue"), ref value);
            Assert.Equal(PropertyMappingResultType.Invalid, result);
            Assert.Equal(1, value);
        }

        [Fact]
        public void WithConverter_NullConverter_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);

            ConvertUsingSimpleMappingDelegate<string> converter = null;
            Assert.Throws<ArgumentNullException>("converter", () => propertyMap.WithConverter(converter));
        }

        [Fact]
        public void WithValueFallback_Invoke_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithValueFallback("abc"));

            FixedValueFallback emptyFallback = Assert.IsType<FixedValueFallback>(propertyMap.EmptyFallback);
            FixedValueFallback invalidFallback = Assert.IsType<FixedValueFallback>(propertyMap.InvalidFallback);

            Assert.Equal("abc", emptyFallback.Value);
            Assert.Equal("abc", invalidFallback.Value);
        }

        [Fact]
        public void WithThrowingFallback_Invoke_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingFallback());

            Assert.IsType<ThrowFallback>(propertyMap.EmptyFallback);
            Assert.IsType<ThrowFallback>(propertyMap.InvalidFallback);
        }

        [Fact]
        public void WithEmptyFallback_Invoke_Success()
        {
            SingleExcelPropertyMap<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithEmptyFallback("abc"));

            FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(mapping.EmptyFallback);
            Assert.Equal("abc", fallback.Value);
        }

        [Fact]
        public void WithThrowingEmptyFallback_Invoke_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingEmptyFallback());

            Assert.IsType<ThrowFallback>(propertyMap.EmptyFallback);
        }

        [Fact]
        public void WithEmptyFallbackItem_ValidFallbackItem_Success()
        {
            IFallbackItem fallback = new FixedValueFallback(10);
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithEmptyFallbackItem(fallback));

            Assert.Same(fallback, propertyMap.EmptyFallback);
        }

        [Fact]
        public void WithEmptyFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("fallbackItem", () => propertyMap.WithEmptyFallbackItem(null));
        }

        [Fact]
        public void WithInvalidFallback_Invoke_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithInvalidFallback("abc"));

            FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(propertyMap.InvalidFallback);
            Assert.Equal("abc", fallback.Value);
        }

        [Fact]
        public void WithThrowingInvalidFallback_Invoke_Success()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingInvalidFallback());

            Assert.IsType<ThrowFallback>(propertyMap.InvalidFallback);
        }

        [Fact]
        public void WithInvalidFallbackItem_ValidFallbackItem_Success()
        {
            IFallbackItem fallback = new FixedValueFallback(10);
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithInvalidFallbackItem(fallback));

            Assert.Same(fallback, propertyMap.InvalidFallback);
        }

        [Fact]
        public void WithInvalidFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
        {
            SingleExcelPropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("fallbackItem", () => propertyMap.WithInvalidFallbackItem(null));
        }
    }
}
