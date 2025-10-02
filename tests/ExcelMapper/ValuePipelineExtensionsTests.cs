using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests;

public class ValuePipelineExtensionsTests : ExcelClassMap<Helpers.TestClass>
{
    [Fact]
    public void WithMapper_ValidMapper_Success()
    {
        var factory = new ColumnNameReaderFactory("ColumnName");
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithReaderFactory(factory));

        Assert.Same(factory, propertyMap.ReaderFactory);
    }

    [Fact]
    public void WithReaderFactory_OptionalColumn_Success()
    {
        var innerReader = new ColumnNameReaderFactory("ColumnName");
        OneToOneMap<string> propertyMap = Map(t => t.Value).MakeOptional();
        Assert.True(propertyMap.Optional);
        Assert.Same(propertyMap, propertyMap.WithReaderFactory(innerReader));
        Assert.True(propertyMap.Optional);
        Assert.Same(innerReader, propertyMap.ReaderFactory);
    }

    [Fact]
    public void WithReaderFactory_NullReader_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("readerFactory", () => propertyMap.WithReaderFactory(null!));
    }

    [Fact]
    public void WithCellValueMappers_ValidMappers_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        ICellMapper mapper1 = Assert.Single(propertyMap.Pipeline.CellValueMappers);
        ICellMapper mapper2 = new BoolMapper();

        Assert.Same(propertyMap, propertyMap.WithCellValueMappers(mapper2));
        Assert.Equal([mapper1, mapper2], propertyMap.Pipeline.CellValueMappers);
    }

    [Fact]
    public void WithCellValueMappers_NullMappers_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("mappers", () => propertyMap.WithCellValueMappers(null!));
    }

    [Fact]
    public void WithCellValueMappers_NullMapperInMapperss_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("mappers", () => propertyMap.WithCellValueMappers([null!]));
    }

    [Fact]
    public void WithMapping_ValidReader_Success()
    {
        var dictionaryMapping = new Dictionary<string, DateTime> { { "key", DateTime.MinValue } };
        StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;

        OneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
        Assert.Same(propertyMap, propertyMap.WithMapping(dictionaryMapping, comparer));

        DictionaryMapper<DateTime> item = propertyMap.Pipeline.CellValueMappers.OfType<DictionaryMapper<DateTime>>().Single();
        Assert.NotSame(dictionaryMapping, item.MappingDictionary);
        Assert.Equal(dictionaryMapping, item.MappingDictionary);

        Assert.Same(comparer, Assert.IsType<Dictionary<string, DateTime>>(item.MappingDictionary).Comparer);
    }

    [Fact]
    public void WithMapping_NullMapping_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("mappingDictionary", () => propertyMap.WithMapping((Dictionary<string, string>)null!));
    }

    [Fact]
    public void MakeOptional_HasMapper_ReturnsExpected()
    {
        var innerReader = new ColumnIndexReaderFactory(1);
        OneToOneMap<string> propertyMap = Map(t => t.Value).WithReaderFactory(innerReader);
        Assert.False(propertyMap.Optional);
        Assert.Same(propertyMap, propertyMap.MakeOptional());
        Assert.True(propertyMap.Optional);
        Assert.Same(innerReader, propertyMap.ReaderFactory);
    }

    [Fact]
    public void MakeOptional_AlreadyOptional_ReturnsExpected()
    {
        var innerReader = new ColumnIndexReaderFactory(1);
        OneToOneMap<string> propertyMap = Map(t => t.Value).WithReaderFactory(innerReader);
        Assert.Same(propertyMap, propertyMap.MakeOptional());
        Assert.True(propertyMap.Optional);
        Assert.Same(propertyMap, propertyMap.MakeOptional());
        Assert.True(propertyMap.Optional);
        Assert.Same(innerReader, propertyMap.ReaderFactory);
    }

    [Fact]
    public void WithTrim_Invoke_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithTrim());

        ICellTransformer transformer = Assert.Single(propertyMap.Pipeline.CellValueTransformers);
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

        OneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
        Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

        DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        OneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
        propertyMap.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

        DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_AutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        OneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
        Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

        DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        OneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);
        propertyMap.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

        DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithDateFormats_NullFormats_ThrowsArgumentNullException()
    {
        OneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);

        Assert.Throws<ArgumentNullException>("formats", () => propertyMap.WithDateFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => propertyMap.WithDateFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithDateFormats_EmptyFormats_ThrowsArgumentException()
    {
        OneToOneMap<DateTime> propertyMap = Map(t => t.DateValue);

        Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats([]));
        Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new List<string>()));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NullableAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        OneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
        Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

        DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NullableNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        OneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
        propertyMap.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(propertyMap, propertyMap.WithDateFormats(formatsArray));

        DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NullableAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        OneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
        Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

        DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NullableNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        OneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);
        propertyMap.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(propertyMap, propertyMap.WithDateFormats(formats));

        DateTimeMapper item = propertyMap.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithDateFormats_NullableNullFormats_ThrowsArgumentNullException()
    {
        OneToOneMap<DateTime?> mapping = Map(t => t.NullableDateValue);

        Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => mapping.WithDateFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithDateFormats_NullableEmptyFormats_ThrowsArgumentException()
    {
        OneToOneMap<DateTime?> propertyMap = Map(t => t.NullableDateValue);

        Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats([]));
        Assert.Throws<ArgumentException>("formats", () => propertyMap.WithDateFormats(new List<string>()));
    }

    [Fact]
    public void WithConverter_SuccessConverterSimple_ReturnsExpected()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithConverter(stringValue =>
        {
            Assert.Equal("stringValue", stringValue);
            return "abc";
        }));
        ConvertUsingMapper item = propertyMap.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

        CellMapperResult result = item.Converter(new ReadCellResult(-1, "stringValue"));
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

        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithConverter(converter));
        ConvertUsingMapper item = propertyMap.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

        CellMapperResult result = item.Converter(new ReadCellResult(-1, "stringValue"));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.IsType<NotSupportedException>(result.Exception);
    }

    [Fact]
    public void WithConverter_SuccessConverterAdvanced_ReturnsExpected()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithConverter(readResult =>
        {
            Assert.Equal("stringValue", readResult.StringValue);
            return CellMapperResult.Success("abc");
        }));
        ConvertUsingMapper item = propertyMap.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

        CellMapperResult result = item.Converter(new ReadCellResult(-1, "stringValue"));
        Assert.True(result.Succeeded);
        Assert.Equal("abc", result.Value);
        Assert.Null(result.Exception);
    }

    [Fact]
    public void WithConverter_InvalidConverterAdvanced_ReturnsExpected()
    {
        ConvertUsingMapperDelegate converter = readResult =>
        {
            Assert.Equal("stringValue", readResult.StringValue);
            throw new NotSupportedException();
        };

        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithConverter(converter));
        ConvertUsingMapper item = propertyMap.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

        Assert.Throws<NotSupportedException>(() => item.Converter(new ReadCellResult(-1, "stringValue")));
    }

    [Fact]
    public void WithConverter_NullConverter_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("converter", () => propertyMap.WithConverter((ConvertUsingSimpleMapperDelegate<string>)null!));
        Assert.Throws<ArgumentNullException>("converter", () => propertyMap.WithConverter((ConvertUsingMapperDelegate)null!));
    }

    [Fact]
    public void WithValueFallback_Invoke_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithValueFallback("abc"));

        FixedValueFallback emptyFallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.EmptyFallback);
        FixedValueFallback invalidFallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.InvalidFallback);

        Assert.Equal("abc", emptyFallback.Value);
        Assert.Equal("abc", invalidFallback.Value);
    }

    [Fact]
    public void WithThrowingFallback_Invoke_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithThrowingFallback());

        Assert.IsType<ThrowFallback>(propertyMap.Pipeline.EmptyFallback);
        Assert.IsType<ThrowFallback>(propertyMap.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithEmptyFallback_Invoke_Success()
    {
        OneToOneMap<string> propertyMapping = Map(t => t.Value);
        Assert.Same(propertyMapping, propertyMapping.WithEmptyFallback("abc"));

        FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(propertyMapping.Pipeline.EmptyFallback);
        Assert.Equal("abc", fallback.Value);
    }

    [Fact]
    public void WithThrowingEmptyFallback_Invoke_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithThrowingEmptyFallback());

        Assert.IsType<ThrowFallback>(propertyMap.Pipeline.EmptyFallback);
    }

    [Fact]
    public void WithEmptyFallbackItem_ValidFallbackItem_Success()
    {
        IFallbackItem fallback = new FixedValueFallback(10);
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithEmptyFallbackItem(fallback));

        Assert.Same(fallback, propertyMap.Pipeline.EmptyFallback);
    }

    [Fact]
    public void WithEmptyFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("fallbackItem", () => propertyMap.WithEmptyFallbackItem(null!));
    }

    [Fact]
    public void WithInvalidFallback_Invoke_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithInvalidFallback("abc"));

        FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(propertyMap.Pipeline.InvalidFallback);
        Assert.Equal("abc", fallback.Value);
    }

    [Fact]
    public void WithThrowingInvalidFallback_Invoke_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithThrowingInvalidFallback());

        Assert.IsType<ThrowFallback>(propertyMap.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithInvalidFallbackItem_ValidFallbackItem_Success()
    {
        IFallbackItem fallback = new FixedValueFallback(10);
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithInvalidFallbackItem(fallback));

        Assert.Same(fallback, propertyMap.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithInvalidFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("fallbackItem", () => propertyMap.WithInvalidFallbackItem(null!));
    }
}
