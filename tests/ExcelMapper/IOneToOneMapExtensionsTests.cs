using System.Linq;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class IOneToOneMapExtensionsTests : ExcelClassMap<Helpers.TestClass>
    {
        [Fact]
        public void WithColumnName_ValidColumnName_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            ColumnNameValueReader reader = Assert.IsType<ColumnNameValueReader>(propertyMap.Reader);
            Assert.Equal("ColumnName", reader.ColumnName);
        }

        [Fact]
        public void WithColumnNameMatching_ValidColumnName_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value).WithColumnNameMatching(e => e == "ColumnName");
            Assert.Same(propertyMap, propertyMap.WithColumnNameMatching(e => e == "ColumnName"));

            Assert.IsType<ColumnNameMatchingValueReader>(propertyMap.Reader);
        }

        [Fact]
        public void WithColumnName_OptionalColumn_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));
            Assert.True(propertyMap.Optional);

            ColumnNameValueReader innerReader = Assert.IsType<ColumnNameValueReader>(propertyMap.Reader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("columnName", () => propertyMap.WithColumnName(null));
        }

        [Fact]
        public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentException>("columnName", () => propertyMap.WithColumnName(string.Empty));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_ValidColumnIndex_Success(int columnIndex)
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

            ColumnIndexValueReader reader = Assert.IsType<ColumnIndexValueReader>(propertyMap.Reader);
            Assert.Equal(columnIndex, reader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_OptionalColumn_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(1));
            Assert.True(propertyMap.Optional);

            ColumnIndexValueReader innerReader = Assert.IsType<ColumnIndexValueReader>(propertyMap.Reader);
            Assert.Equal(1, innerReader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => propertyMap.WithColumnIndex(-1));
        }
        
        [Fact]
        public void WithMapper_ValidMapper_Success()
        {
            var reader = new ColumnNameValueReader("ColumnName");
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithReader(reader));

            Assert.Same(reader, propertyMap.Reader);
        }

        [Fact]
        public void WithReader_OptionalColumn_Success()
        {
            var innerReader = new ColumnNameValueReader("ColumnName");
            IOneToOneMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.WithReader(innerReader));
            Assert.True(propertyMap.Optional);
            Assert.Same(innerReader, propertyMap.Reader);
        }

        [Fact]
        public void WithReader_NullReader_ThrowsArgumentNullException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("reader", () => propertyMap.WithReader(null));
        }

        [Fact]
        public void WithCellValueMappers_ValidMappersSingle_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            var expected = propertyMap.Mappers.ToList();
            ICellValueMapper mapper = new BoolMapper();

            expected.Insert(expected.Count - 1, mapper);
            Assert.Same(propertyMap, propertyMap.WithCellValueMappers(mapper));
            Assert.Equal(expected, propertyMap.Mappers.Select(c => c));
        }

        [Fact]
        public void WithCellValueMappers_ValidMappersMultiple_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            var expected = propertyMap.Mappers.ToList();
            ICellValueMapper mapper1 = new BoolMapper();
            ICellValueMapper mapper2 = new GuidMapper();

            expected.Insert(expected.Count - 1, mapper1);
            expected.Insert(expected.Count - 1, mapper2);
            Assert.Same(propertyMap, propertyMap.WithCellValueMappers(mapper1, mapper2));
            Assert.Equal(expected, propertyMap.Mappers);
        }

        [Fact]
        public void WithCellValueMappers_NullMappers_ThrowsArgumentNullException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("mappers", () => propertyMap.WithCellValueMappers(null));
        }

        [Fact]
        public void WithCellValueMappers_NullMapperInMapperss_ThrowsArgumentNullException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("mappers", () => propertyMap.WithCellValueMappers(new ICellValueMapper[] { null }));
        }

        [Fact]
        public void WithMapping_ValidReader_Success()
        {
            var dictionaryMapping = new Dictionary<string, DateTime> { { "key", DateTime.MinValue } };
            StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;

            IOneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithMapping(dictionaryMapping, comparer));

            DictionaryMapper<DateTime> item = propertyMap.Mappers.OfType<DictionaryMapper<DateTime>>().Single();
            Assert.NotSame(dictionaryMapping, item.MappingDictionary);
            Assert.Equal(dictionaryMapping, item.MappingDictionary);

            Assert.Same(comparer, Assert.IsType<Dictionary<string, DateTime>>(item.MappingDictionary).Comparer);
        }

        [Fact]
        public void WithMapping_NullMapping_ThrowsArgumentNullException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("mappingDictionary", () => propertyMap.WithMapping((Dictionary<string, string>)null));
        }

        [Fact]
        public void MakeOptional_HasMapper_ReturnsExpected()
        {
            var innerReader = new ColumnIndexValueReader(1);
            IOneToOneMap<string> propertyMap = Map(t => t.Value).WithReader(innerReader);
            Assert.False(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(innerReader, propertyMap.Reader);
        }

        [Fact]
        public void MakeOptional_AlreadyOptional_ReturnsExpected()
        {
            var innerReader = new ColumnIndexValueReader(1);
            IOneToOneMap<string> propertyMap = Map(t => t.Value).WithReader(innerReader);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(innerReader, propertyMap.Reader);
        }

        [Fact]
        public void WithTrim_Invoke_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithTrim());

            Assert.IsType<TrimCellValueTransformer>(propertyMap.Mappers[0]);
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

            IOneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.Mappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NotMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            IOneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
            propertyMap.Mappers.RemoveAt(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.Mappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_AutoMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            IOneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.Mappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NotMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            IOneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
            propertyMap.Mappers.RemoveAt(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.Mappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Fact]
        public void WithDateFormats_NullFormats_ThrowsArgumentNullException()
        {
            IOneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);

            Assert.Throws<ArgumentNullException>("formats", () => propertyMap.WithDateFormats(null));
            Assert.Throws<ArgumentNullException>("formats", () => propertyMap.WithDateFormats((IEnumerable<string>)null));
        }

        [Fact]
        public void WithDateFormats_EmptyFormats_ThrowsArgumentException()
        {
            IOneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);

            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new string[0]));
            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new List<string>()));
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableAutoMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            IOneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.Mappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableNotMappedStringArray_Success(IEnumerable<string> formats)
        {
            var formatsArray = formats.ToArray();

            IOneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            propertyMap.Mappers.RemoveAt(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

            DateTimeMapper item = propertyMap.Mappers.OfType<DateTimeMapper>().Single();
            Assert.Same(formatsArray, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableAutoMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            IOneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.Mappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Theory]
        [MemberData(nameof(Formats_TestData))]
        public void WithDateFormats_NullableNotMappedIEnumerableString_Success(IEnumerable<string> formats)
        {
            IOneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
            propertyMap.Mappers.RemoveAt(0);

            Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

            DateTimeMapper item = propertyMap.Mappers.OfType<DateTimeMapper>().Single();
            Assert.Equal(formats, item.Formats);
        }

        [Fact]
        public void WithDateFormats_NullableNullFormats_ThrowsArgumentNullException()
        {
            IOneToOneMap<DateTime?> mapping = Map(t => t.NullableDateValue);

            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats(null));
            Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats((IEnumerable<string>)null));
        }

        [Fact]
        public void WithDateFormats_NullableEmptyFormats_ThrowsArgumentException()
        {
            IOneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);

            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new string[0]));
            Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new List<string>()));
        }

        [Fact]
        public void WithConverter_SuccessConverterSimple_ReturnsExpected()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithConverter(stringValue =>
            {
                Assert.Equal("stringValue", stringValue);
                return "abc";
            }));
            ConvertUsingMapper item = propertyMap.Mappers.OfType<ConvertUsingMapper>().Single();

            CellValueMapperResult result = item.Converter(default, new CellValueMapperResult("stringValue", null, CellValueMapperResult.HandleAction.UseResultAndContinueMapping));
            Assert.True(result.Succeeded);
            Assert.Equal("abc", result.Value);
            Assert.Null(result.Exception);
        }

        [Fact]
        public void WithConverter_InvalidConverter_ReturnsExpected()
        {
            ConvertUsingSimpleMapperDelegate<string> converter = stringValue =>
            {
                Assert.Equal("stringValue", stringValue);
                throw new NotSupportedException();
            };

            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithConverter(converter));
            ConvertUsingMapper item = propertyMap.Mappers.OfType<ConvertUsingMapper>().Single();

            CellValueMapperResult result = item.Converter(default, new CellValueMapperResult("stringValue", null, CellValueMapperResult.HandleAction.UseResultAndContinueMapping));
            Assert.False(result.Succeeded);
            Assert.Equal("stringValue", result.Value);
            Assert.IsType<NotSupportedException>(result.Exception);
        }

        [Fact]
        public void WithConverter_SuccessConverterAdvanced_ReturnsExpected()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            var cell = new ExcelCell(null, 1, 2);
            Assert.Same(propertyMap, propertyMap.WithConverter((cell, previous) =>
            {
                Assert.Equal(1, cell.RowIndex);
                Assert.Equal(2, cell.ColumnIndex);
                Assert.Equal("stringValue", previous.Value);
                return previous.Success("abc");
            }));
            ConvertUsingMapper item = propertyMap.Mappers.OfType<ConvertUsingMapper>().Single();

            CellValueMapperResult result = item.Converter(cell, new CellValueMapperResult("stringValue", null, CellValueMapperResult.HandleAction.UseResultAndContinueMapping));
            Assert.True(result.Succeeded);
            Assert.Equal("abc", result.Value);
            Assert.Null(result.Exception);
        }

        [Fact]
        public void WithConverter_InvalidConverterAdvanced_ReturnsExpected()
        {
            var cell = new ExcelCell(null, 1, 2);
            ConvertUsingMapperDelegate converter = (cell, previous) =>
            {
                Assert.Equal(1, cell.RowIndex);
                Assert.Equal(2, cell.ColumnIndex);
                Assert.Equal("stringValue", previous.Value);
                throw new NotSupportedException();
            };

            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithConverter(converter));
            ConvertUsingMapper item = propertyMap.Mappers.OfType<ConvertUsingMapper>().Single();

            Assert.Throws<NotSupportedException>(() => item.Converter(cell, new CellValueMapperResult("stringValue", null, CellValueMapperResult.HandleAction.UseResultAndContinueMapping)));
        }

        [Fact]
        public void WithConverter_NullConverter_ThrowsArgumentNullException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("converter", () => propertyMap.WithConverter((ConvertUsingSimpleMapperDelegate<string>)null));
            Assert.Throws<ArgumentNullException>("converter", () => propertyMap.WithConverter((ConvertUsingMapperDelegate)null));
        }

#if TODO_NEW
        [Fact]
        public void WithValueFallback_Invoke_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithValueFallback("abc"));

            FixedValueFallback emptyFallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.EmptyFallback);
            FixedValueFallback invalidFallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.InvalidFallback);

            Assert.Equal("abc", emptyFallback.Value);
            Assert.Equal("abc", invalidFallback.Value);
        }

        [Fact]
        public void WithThrowingFallback_Invoke_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingFallback());

            Assert.IsType<ThrowFallback>(propertyMap.Pipeline.EmptyFallback);
            Assert.IsType<ThrowFallback>(propertyMap.Pipeline.InvalidFallback);
        }

        [Fact]
        public void WithEmptyFallback_Invoke_Success()
        {
            IOneToOneMap<string> propertyMapping = Map(t => t.Value);
            Assert.Same(propertyMapping, propertyMapping.WithEmptyFallback("abc"));

            FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(propertyMapping.Pipeline.EmptyFallback);
            Assert.Equal("abc", fallback.Value);
        }

        [Fact]
        public void WithThrowingEmptyFallback_Invoke_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingEmptyFallback());

            Assert.IsType<ThrowFallback>(propertyMap.Pipeline.EmptyFallback);
        }

        [Fact]
        public void WithEmptyFallbackItem_ValidFallbackItem_Success()
        {
            IFallbackItem fallback = new FixedValueFallback(10);
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithEmptyFallbackItem(fallback));

            Assert.Same(fallback, propertyMap.Pipeline.EmptyFallback);
        }

        [Fact]
        public void WithEmptyFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentNullException>("fallbackItem", () => propertyMap.WithEmptyFallbackItem(null));
        }

        [Fact]
        public void WithInvalidFallback_Invoke_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithInvalidFallback("abc"));

            FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.InvalidFallback);
            Assert.Equal("abc", fallback.Value);
        }

        [Fact]
        public void WithThrowingInvalidFallback_Invoke_Success()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithThrowingInvalidFallback());

            Assert.IsType<ThrowFallback>(propertyMap.Pipeline.InvalidFallback);
        }

        [Fact]
        public void WithInvalidFallbackItem_ValidFallbackItem_Success()
        {
            IFallbackItem fallback = new FixedValueFallback(10);
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithInvalidFallbackItem(fallback));

            Assert.Same(fallback, propertyMap.Pipeline.InvalidFallback);
        }

        [Fact]
        public void WithInvalidFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
        {
            IOneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("fallbackItem", () => propertyMap.WithInvalidFallbackItem(null));
        }
#endif
    }
}
