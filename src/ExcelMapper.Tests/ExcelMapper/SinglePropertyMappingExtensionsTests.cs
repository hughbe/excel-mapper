using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Items;
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

            ColumnPropertyMapper mapper = Assert.IsType<ColumnPropertyMapper>(mapping.Mapper);
            Assert.Equal("ColumnName", mapper.ColumnName);
        }

        [Fact]
        public void WithColumnName_OptionalColumn_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value).MakeOptional();
            Assert.Same(mapping, mapping.WithColumnName("ColumnName"));

            OptionalPropertyMapper mapper = Assert.IsType<OptionalPropertyMapper>(mapping.Mapper);
            ColumnPropertyMapper innerMapper = Assert.IsType<ColumnPropertyMapper>(mapper.Mapper);
            Assert.Equal("ColumnName", innerMapper.ColumnName);
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
        public void WithIndex_ValidColumnIndex_Success(int columnIndex)
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithIndex(columnIndex));

            IndexPropertyMapper mapper = Assert.IsType<IndexPropertyMapper>(mapping.Mapper);
            Assert.Equal(columnIndex, mapper.ColumnIndex);
        }

        [Fact]
        public void WithIndex_OptionalColumn_Success()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value).MakeOptional();
            Assert.Same(mapping, mapping.WithIndex(1));

            OptionalPropertyMapper mapper = Assert.IsType<OptionalPropertyMapper>(mapping.Mapper);
            IndexPropertyMapper innerMapper = Assert.IsType<IndexPropertyMapper>(mapper.Mapper);
            Assert.Equal(1, innerMapper.ColumnIndex);
        }

        [Fact]
        public void WithIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => mapping.WithIndex(-1));
        }

        [Fact]
        public void WithMapper_ValidMapper_Success()
        {
            var mapper = new ColumnPropertyMapper("ColumnName");
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Same(mapping, mapping.WithMapper(mapper));

            Assert.Same(mapper, mapping.Mapper);
        }

        [Fact]
        public void WithMapping_OptionalColumn_Success()
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");

            SinglePropertyMapping<string> mapping = Map(t => t.Value).MakeOptional();
            Assert.Same(mapping, mapping.WithMapper(innerMapper));

            OptionalPropertyMapper mapper = Assert.IsType<OptionalPropertyMapper>(mapping.Mapper);
            Assert.Same(innerMapper, mapper.Mapper);
        }

        [Fact]
        public void WithMapping_ValidMapper_Success()
        {
            var dictionaryMapping = new Dictionary<string, DateTime> { { "key", DateTime.MinValue } };
            StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;

            SinglePropertyMapping<DateTime> mapping = Map(t => t.DateValue);
            Assert.Same(mapping, mapping.WithMapping(dictionaryMapping, comparer));

            MapStringValueMappingItem<DateTime> item = mapping.MappingItems.OfType<MapStringValueMappingItem<DateTime>>().Single();
            Assert.NotSame(dictionaryMapping, item.Mapping);
            Assert.Equal(dictionaryMapping, item.Mapping);

            Assert.Same(comparer, Assert.IsType<Dictionary<string, DateTime>>(item.Mapping).Comparer);
        }

        [Fact]
        public void WithMapping_NullMapper_ThrowsArgumentNullException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("mapper", () => mapping.WithMapper(null));
        }

        [Fact]
        public void MakeOptional_HasMapper_ReturnsExpected()
        {
            var indexMapper = new IndexPropertyMapper(1);
            SinglePropertyMapping<string> mapping = Map(t => t.Value).WithMapper(indexMapper);
            Assert.Same(mapping, mapping.MakeOptional());

            OptionalPropertyMapper mapper = Assert.IsType<OptionalPropertyMapper>(mapping.Mapper);
            Assert.Same(indexMapper, mapper.Mapper);
        }

        [Fact]
        public void MakeOptional_AlreadyOptional_ThrowsExcelMappingException()
        {
            SinglePropertyMapping<string> mapping = Map(t => t.Value).WithIndex(1).MakeOptional();

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

            PropertyMappingResult result = item.Converter(new MapResult(-1, "stringValue"));
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

            PropertyMappingResult result = item.Converter(new MapResult(-1, "stringValue"));
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
            ISinglePropertyMappingItem fallback = new FixedValueFallback(10);
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
            ISinglePropertyMappingItem fallback = new FixedValueFallback(10);
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
