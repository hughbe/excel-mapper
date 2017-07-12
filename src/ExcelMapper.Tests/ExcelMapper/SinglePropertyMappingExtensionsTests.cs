using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Items;
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
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithColumnName("ColumnName"));

            ColumnNameReader reader = Assert.IsType<ColumnNameReader>(mapping.Reader);
            Assert.Equal("ColumnName", reader.ColumnName);
        }

        [Fact]
        public void WithColumnName_OptionalColumn_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value).MakeOptional();
            Assert.Same(mapping, mapping.WithColumnName("ColumnName"));

            OptionalColumnReader reader = Assert.IsType<OptionalColumnReader>(mapping.Reader);
            ColumnNameReader innerReader = Assert.IsType<ColumnNameReader>(reader.InnerReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("columnName", () => mapping.WithColumnName(null));
        }

        [Fact]
        public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Throws<ArgumentException>("columnName", () => mapping.WithColumnName(string.Empty));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_ValidColumnIndex_Success(int columnIndex)
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithColumnIndex(columnIndex));

            ColumnIndexReader reader = Assert.IsType<ColumnIndexReader>(mapping.Reader);
            Assert.Equal(columnIndex, reader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_OptionalColumn_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value).MakeOptional();
            Assert.Same(mapping, mapping.WithColumnIndex(1));

            OptionalColumnReader reader = Assert.IsType<OptionalColumnReader>(mapping.Reader);
            ColumnIndexReader innerReader = Assert.IsType<ColumnIndexReader>(reader.InnerReader);
            Assert.Equal(1, innerReader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => mapping.WithColumnIndex(-1));
        }

        [Fact]
        public void WithMapper_ValidMapper_Success()
        {
            var reader = new ColumnNameReader("ColumnName");
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithReader(reader));

            Assert.Same(reader, mapping.Reader);
        }

        [Fact]
        public void WithReader_OptionalColumn_Success()
        {
            var innerReader = new ColumnNameReader("ColumnName");

            SinglePropertyMapping<string> mapping = Map(t => t.Value).MakeOptional();
            Assert.Same(mapping, mapping.WithReader(innerReader));

            OptionalColumnReader reader = Assert.IsType<OptionalColumnReader>(mapping.Reader);
            Assert.Same(innerReader, reader.InnerReader);
        }

        [Fact]
        public void WithReader_NullReader_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("reader", () => mapping.WithReader(null));
        }

        [Fact]
        public void WithMapping_ValidReader_Success()
        {
            var dictionaryMapping = new Dictionary<string, DateTime> { { "key", DateTime.MinValue } };
            StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;

            SinglePropertyMapping<DateTime> mapping = Map(t => t.DateValue);
            Assert.Same(mapping, mapping.WithMapping(dictionaryMapping, comparer));

            MapStringValueMappingItem<DateTime> item = mapping.MappingItems.OfType<MapStringValueMappingItem<DateTime>>().Single();
            Assert.NotSame(dictionaryMapping, item.MappingDictionary);
            Assert.Equal(dictionaryMapping, item.MappingDictionary);

            Assert.Same(comparer, Assert.IsType<Dictionary<string, DateTime>>(item.MappingDictionary).Comparer);
        }

        [Fact]
        public void WithMapping_NullMapping_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("mappingDictionary", () => mapping.WithMapping((Dictionary<string, string>)null));
        }

        [Fact]
        public void MakeOptional_HasMapper_ReturnsExpected()
        {
            var innerReader = new ColumnIndexReader(1);
            SinglePropertyMapping<string> mapping = Map(t => t.Value).WithReader(innerReader);
            Assert.Same(mapping, mapping.MakeOptional());

            OptionalColumnReader reader = Assert.IsType<OptionalColumnReader>(mapping.Reader);
            Assert.Same(innerReader, reader.InnerReader);
        }

        [Fact]
        public void MakeOptional_AlreadyOptional_ThrowsExcelMappingException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value).WithColumnIndex(1).MakeOptional();

            Assert.Throws<ExcelMappingException>(() => mapping.MakeOptional());
        }

        [Fact]
        public void WithTrim_Invoke_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithTrim());

            IStringValueTransformer transformer = Assert.Single(mapping.StringValueTransformers);
            Assert.IsType<TrimStringTransformer>(transformer);
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

            SinglePropertyMapping<DateTime> mapping = Map(t => t.DateValue);
            Assert.Same(mapping, mapping.WithDateFormats(formatsArray));

            ParseAsDateTimeMappingItem item = mapping.MappingItems.OfType<ParseAsDateTimeMappingItem>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NotMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            SinglePropertyMapping<DateTime> mapping = Map(t => t.DateValue);
            mapping.RemoveMappingItem(0);

            Assert.Same(mapping, mapping.WithDateFormats(formatsArray));

            ParseAsDateTimeMappingItem item = mapping.MappingItems.OfType<ParseAsDateTimeMappingItem>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_AutoMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            SinglePropertyMapping<DateTime> mapping = Map(t => t.DateValue);
            Assert.Same(mapping, mapping.WithDateFormats(formats));

            ParseAsDateTimeMappingItem item = mapping.MappingItems.OfType<ParseAsDateTimeMappingItem>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NotMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            SinglePropertyMapping<DateTime> mapping = Map(t => t.DateValue);
            mapping.RemoveMappingItem(0);

            Assert.Same(mapping, mapping.WithDateFormats(formats));

            ParseAsDateTimeMappingItem item = mapping.MappingItems.OfType<ParseAsDateTimeMappingItem>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Fact]
        public void WithDateFormats_NullFormats_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<DateTime> mapping = Map(t => t.DateValue);

            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats(null));
            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats((IEnumerable<string>)null));
        }

        [Fact]
        public void WithDateFormats_EmptyFormats_ThrowsArgumentException()
        {
            SinglePropertyMapping<DateTime> mapping = Map(t => t.DateValue);

            Assert.Throws<ArgumentException>("formats", () => mapping.WithDateFormats(new string[0]));
            Assert.Throws<ArgumentException>("formats", () => mapping.WithDateFormats(new List<string>()));
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableAutoMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            SinglePropertyMapping<DateTime?> mapping = Map(t => t.NullableDateValue);
            Assert.Same(mapping, mapping.WithDateFormats(formatsArray));

            ParseAsDateTimeMappingItem item = mapping.MappingItems.OfType<ParseAsDateTimeMappingItem>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableNotMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            SinglePropertyMapping<DateTime?> mapping = Map(t => t.NullableDateValue);
            mapping.RemoveMappingItem(0);

            Assert.Same(mapping, mapping.WithDateFormats(formatsArray));

            ParseAsDateTimeMappingItem item = mapping.MappingItems.OfType<ParseAsDateTimeMappingItem>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableAutoMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            SinglePropertyMapping<DateTime?> mapping = Map(t => t.NullableDateValue);
            Assert.Same(mapping, mapping.WithDateFormats(formats));

            ParseAsDateTimeMappingItem item = mapping.MappingItems.OfType<ParseAsDateTimeMappingItem>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableNotMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            SinglePropertyMapping<DateTime?> mapping = Map(t => t.NullableDateValue);
            mapping.RemoveMappingItem(0);

            Assert.Same(mapping, mapping.WithDateFormats(formats));

            ParseAsDateTimeMappingItem item = mapping.MappingItems.OfType<ParseAsDateTimeMappingItem>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Fact]
        public void WithDateFormats_NullableNullFormats_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<DateTime?> mapping = Map(t => t.NullableDateValue);

            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats(null));
            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats((IEnumerable<string>)null));
        }

        [Fact]
        public void WithDateFormats_NullableEmptyFormats_ThrowsArgumentException()
        {
            SinglePropertyMapping<DateTime?> mapping = Map(t => t.NullableDateValue);

            Assert.Throws<ArgumentException>("formats", () => mapping.WithDateFormats(new string[0]));
            Assert.Throws<ArgumentException>("formats", () => mapping.WithDateFormats(new List<string>()));
        }

        [Fact]
        public void WithConverter_SuccessConverter_ReturnsExpected()
        {
            ConvertUsingSimpleMappingDelegate<string> converter = stringValue =>
            {
                Assert.Equal("stringValue", stringValue);
                return "abc";
            };

            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithConverter(converter));
            ConvertUsingMappingItem item = mapping.MappingItems.OfType<ConvertUsingMappingItem>().Single();

            PropertyMappingResult result = item.Converter(new ReadResult(-1, "stringValue"));
            Assert.Equal(PropertyMappingResultType.Success, result.Type);
            Assert.Equal("abc", result.Value);
        }

        [Fact]
        public void WithConverter_InvalidConverter_ReturnsExpected()
        {
            ConvertUsingSimpleMappingDelegate<string> converter = stringValue =>
            {
                Assert.Equal("stringValue", stringValue);
                throw new NotSupportedException();
            };

            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithConverter(converter));
            ConvertUsingMappingItem item = mapping.MappingItems.OfType<ConvertUsingMappingItem>().Single();

            PropertyMappingResult result = item.Converter(new ReadResult(-1, "stringValue"));
            Assert.Equal(PropertyMappingResultType.Invalid, result.Type);
            Assert.Null(result.Value);
        }

        [Fact]
        public void WithConverter_NullConverter_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);

            ConvertUsingSimpleMappingDelegate<string> converter = null;
            Assert.Throws<ArgumentNullException>("converter", () => mapping.WithConverter(converter));
        }

        [Fact]
        public void WithValueFallback_Invoke_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithValueFallback("abc"));

            FixedValueFallback emptyFallback = Assert.IsType<FixedValueFallback>(mapping.EmptyFallback);
            FixedValueFallback invalidFallback = Assert.IsType<FixedValueFallback>(mapping.InvalidFallback);

            Assert.Equal("abc", emptyFallback.Value);
            Assert.Equal("abc", invalidFallback.Value);
        }

        [Fact]
        public void WithThrowingFallback_Invoke_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithThrowingFallback());

            Assert.IsType<ThrowFallback>(mapping.EmptyFallback);
            Assert.IsType<ThrowFallback>(mapping.InvalidFallback);
        }

        [Fact]
        public void WithEmptyFallback_Invoke_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithEmptyFallback("abc"));

            FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(mapping.EmptyFallback);
            Assert.Equal("abc", fallback.Value);
        }

        [Fact]
        public void WithThrowingEmptyFallback_Invoke_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithThrowingEmptyFallback());

            Assert.IsType<ThrowFallback>(mapping.EmptyFallback);
        }

        [Fact]
        public void WithEmptyFallbackItem_ValidFallbackItem_Success()
        {
            IFallbackItem fallback = new FixedValueFallback(10);
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithEmptyFallbackItem(fallback));

            Assert.Same(fallback, mapping.EmptyFallback);
        }

        [Fact]
        public void WithEmptyFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("fallbackItem", () => mapping.WithEmptyFallbackItem(null));
        }

        [Fact]
        public void WithInvalidFallback_Invoke_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithInvalidFallback("abc"));

            FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(mapping.InvalidFallback);
            Assert.Equal("abc", fallback.Value);
        }

        [Fact]
        public void WithThrowingInvalidFallback_Invoke_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithThrowingInvalidFallback());

            Assert.IsType<ThrowFallback>(mapping.InvalidFallback);
        }

        [Fact]
        public void WithInvalidFallbackItem_ValidFallbackItem_Success()
        {
            IFallbackItem fallback = new FixedValueFallback(10);
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithInvalidFallbackItem(fallback));

            Assert.Same(fallback, mapping.InvalidFallback);
        }

        [Fact]
        public void WithInvalidFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("fallbackItem", () => mapping.WithInvalidFallbackItem(null));
        }
    }
}
