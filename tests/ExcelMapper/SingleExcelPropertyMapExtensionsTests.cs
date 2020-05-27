using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class SingleExcelPropertyMapExtensionsTests : ExcelClassMap<Helpers.TestClass>
    {
        [Fact]
        public void WithColumnName_ValidColumnName_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            ColumnNameValueReader reader = Assert.IsType<ColumnNameValueReader>(propertyMap.CellReader);
            Assert.Equal("ColumnName", reader.ColumnName);
        }

        [Fact]
        public void WithColumnNameMatching_ValidColumnName_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value).WithColumnNameMatching(e => e == "ColumnName");
            Assert.Same(propertyMap, propertyMap.WithColumnNameMatching(e => e == "ColumnName"));

            Assert.IsType<ColumnNameMatchingValueReader>(propertyMap.CellReader);
        }

        [Fact]
        public void WithColumnName_OptionalColumn_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));
            Assert.True(propertyMap.Optional);

            ColumnNameValueReader innerReader = Assert.IsType<ColumnNameValueReader>(propertyMap.CellReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("columnName", () => propertyMap.WithColumnName(null));
        }

        [Fact]
        public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentException>("columnName", () => propertyMap.WithColumnName(string.Empty));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_ValidColumnIndex_Success(int columnIndex)
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

            ColumnIndexValueReader reader = Assert.IsType<ColumnIndexValueReader>(propertyMap.CellReader);
            Assert.Equal(columnIndex, reader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_OptionalColumn_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(1));
            Assert.True(propertyMap.Optional);

            ColumnIndexValueReader innerReader = Assert.IsType<ColumnIndexValueReader>(propertyMap.CellReader);
            Assert.Equal(1, innerReader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => propertyMap.WithColumnIndex(-1));
        }

        [Fact]
        public void WithMapper_ValidMapper_Success()
        {
            var reader = new ColumnNameValueReader("ColumnName");
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithReader(reader));

            Assert.Same(reader, propertyMap.CellReader);
        }

        [Fact]
        public void WithReader_OptionalColumn_Success()
        {
            var innerReader = new ColumnNameValueReader("ColumnName");
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.WithReader(innerReader));
            Assert.True(propertyMap.Optional);
            Assert.Same(innerReader, propertyMap.CellReader);
        }

        [Fact]
        public void WithReader_NullReader_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("reader", () => propertyMap.WithReader(null));
        }

        [Fact]
        public void WithCellValueMappers_ValidMappers_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            ICellValueMapper mapper1 = Assert.Single(propertyMap.Pipeline.CellValueMappers);
            ICellValueMapper mapper2 = new BoolMapper(); ;

            Assert.Same(propertyMap, propertyMap.WithCellValueMappers(mapper2));
            Assert.Equal(new ICellValueMapper[] { mapper1, mapper2 }, propertyMap.Pipeline.CellValueMappers);
        }

        [Fact]
        public void WithCellValueMappers_NullMappers_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("mappers", () => propertyMap.WithCellValueMappers(null));
        }

        [Fact]
        public void WithCellValueMappers_NullMapperInMapperss_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("mappers", () => propertyMap.WithCellValueMappers(new ICellValueMapper[] { null }));
        }

        [Fact]
        public void WithMapping_ValidReader_Success()
        {
            var dictionaryMapping = new Dictionary<string, DateTime> { { "key", DateTime.MinValue } };
            StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;

            OneToOnePropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithMapping(dictionaryMapping, comparer));

            DictionaryMapper<DateTime> item = propertyMap.Pipeline.CellValueMappers.OfType<DictionaryMapper<DateTime>>().Single();
            Assert.NotSame(dictionaryMapping, item.MappingDictionary);
            Assert.Equal(dictionaryMapping, item.MappingDictionary);

            Assert.Same(comparer, Assert.IsType<Dictionary<string, DateTime>>(item.MappingDictionary).Comparer);
        }

        [Fact]
        public void WithMapping_NullMapping_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("mappingDictionary", () => propertyMap.WithMapping((Dictionary<string, string>)null));
        }

        [Fact]
        public void MakeOptional_HasMapper_ReturnsExpected()
        {
            var innerReader = new ColumnIndexValueReader(1);
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value).WithReader(innerReader);
            Assert.False(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(innerReader, propertyMap.CellReader);
        }

        [Fact]
        public void MakeOptional_AlreadyOptional_ReturnsExpected()
        {
            var innerReader = new ColumnIndexValueReader(1);
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value).WithReader(innerReader);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(innerReader, propertyMap.CellReader);
        }

        [Fact]
        public void WithTrim_Invoke_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithTrim());

            ICellValueTransformer transformer = Assert.Single(propertyMap.Pipeline.CellValueTransformers);
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

            OneToOnePropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NotMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            OneToOnePropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            propertyMap.Pipeline.RemoveCellValueMapper(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_AutoMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            OneToOnePropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NotMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            OneToOnePropertyMap<DateTime> propertyMap = Map(t => t.DateValue);
            propertyMap.Pipeline.RemoveCellValueMapper(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Fact]
        public void WithDateFormats_NullFormats_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<DateTime> propertyMap = Map(t => t.DateValue);

            Assert.Throws<ArgumentNullException>("formats", () => propertyMap.WithDateFormats(null));
            Assert.Throws<ArgumentNullException>("formats", () => propertyMap.WithDateFormats((IEnumerable<string>)null));
        }

        [Fact]
        public void WithDateFormats_EmptyFormats_ThrowsArgumentException()
        {
            OneToOnePropertyMap<DateTime> propertyMap = Map(t => t.DateValue);

            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new string[0]));
            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new List<string>()));
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableAutoMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            OneToOnePropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableNotMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            OneToOnePropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            propertyMap.Pipeline.RemoveCellValueMapper(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableAutoMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            OneToOnePropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableNotMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            OneToOnePropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            propertyMap.Pipeline.RemoveCellValueMapper(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Fact]
        public void WithDateFormats_NullableNullFormats_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<DateTime?> mapping = Map(t => t.NullableDateValue);

            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats(null));
            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats((IEnumerable<string>)null));
        }

        [Fact]
        public void WithDateFormats_NullableEmptyFormats_ThrowsArgumentException()
        {
            OneToOnePropertyMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);

            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new string[0]));
            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new List<string>()));
        }

        [Fact]
        public void WithConverter_SuccessConverter_ReturnsExpected()
        {
            ConvertUsingSimpleMapperDelegate<string> converter = stringValue =>
            {
                Assert.Equal("stringValue", stringValue);
                return "abc";
            };

            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithConverter(converter));
            ConvertUsingMapper item = propertyMap.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

            object value = null;
            PropertyMapperResultType result = item.Converter(new ReadCellValueResult(-1, "stringValue"), ref value);
            Assert.Equal(PropertyMapperResultType.Success, result);
            Assert.Equal("abc", value);
        }

        [Fact]
        public void WithConverter_InvalidConverter_ReturnsExpected()
        {
            ConvertUsingSimpleMapperDelegate<string> converter = stringValue =>
            {
                Assert.Equal("stringValue", stringValue);
                throw new NotSupportedException();
            };

            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithConverter(converter));
            ConvertUsingMapper item = propertyMap.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

            object value = 1;
            PropertyMapperResultType result = item.Converter(new ReadCellValueResult(-1, "stringValue"), ref value);
            Assert.Equal(PropertyMapperResultType.Invalid, result);
            Assert.Equal(1, value);
        }

        [Fact]
        public void WithConverter_NullConverter_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);

            ConvertUsingSimpleMapperDelegate<string> converter = null;
            Assert.Throws<ArgumentNullException>("converter", () => propertyMap.WithConverter(converter));
        }

        [Fact]
        public void WithValueFallback_Invoke_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithValueFallback("abc"));

            FixedValueFallback emptyFallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.EmptyFallback);
            FixedValueFallback invalidFallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.InvalidFallback);

            Assert.Equal("abc", emptyFallback.Value);
            Assert.Equal("abc", invalidFallback.Value);
        }

        [Fact]
        public void WithThrowingFallback_Invoke_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingFallback());

            Assert.IsType<ThrowFallback>(propertyMap.Pipeline.EmptyFallback);
            Assert.IsType<ThrowFallback>(propertyMap.Pipeline.InvalidFallback);
        }

        [Fact]
        public void WithEmptyFallback_Invoke_Success()
        {
            OneToOnePropertyMap<string> propertyMapping = Map(t => t.Value);
            Assert.Same(propertyMapping, propertyMapping.WithEmptyFallback("abc"));

            FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(propertyMapping.Pipeline.EmptyFallback);
            Assert.Equal("abc", fallback.Value);
        }

        [Fact]
        public void WithThrowingEmptyFallback_Invoke_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingEmptyFallback());

            Assert.IsType<ThrowFallback>(propertyMap.Pipeline.EmptyFallback);
        }

        [Fact]
        public void WithEmptyFallbackItem_ValidFallbackItem_Success()
        {
            IFallbackItem fallback = new FixedValueFallback(10);
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithEmptyFallbackItem(fallback));

            Assert.Same(fallback, propertyMap.Pipeline.EmptyFallback);
        }

        [Fact]
        public void WithEmptyFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("fallbackItem", () => propertyMap.WithEmptyFallbackItem(null));
        }

        [Fact]
        public void WithInvalidFallback_Invoke_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithInvalidFallback("abc"));

            FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.InvalidFallback);
            Assert.Equal("abc", fallback.Value);
        }

        [Fact]
        public void WithThrowingInvalidFallback_Invoke_Success()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingInvalidFallback());

            Assert.IsType<ThrowFallback>(propertyMap.Pipeline.InvalidFallback);
        }

        [Fact]
        public void WithInvalidFallbackItem_ValidFallbackItem_Success()
        {
            IFallbackItem fallback = new FixedValueFallback(10);
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithInvalidFallbackItem(fallback));

            Assert.Same(fallback, propertyMap.Pipeline.InvalidFallback);
        }

        [Fact]
        public void WithInvalidFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
        {
            OneToOnePropertyMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("fallbackItem", () => propertyMap.WithInvalidFallbackItem(null));
        }
    }
}
