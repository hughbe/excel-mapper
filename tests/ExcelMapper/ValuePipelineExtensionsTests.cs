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
    public void WithReaderFactory_ValidReader_Success()
    {
        var factory = new ColumnNameReaderFactory("ColumnName");
        var map = Map(t => t.Value);
        Assert.False(map.Optional);
        Assert.Same(map, map.WithReaderFactory(factory));
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void WithReaderFactory_OptionalColumn_Success()
    {
        var factory = new ColumnNameReaderFactory("ColumnName");
        var map = Map(t => t.Value).MakeOptional();
        Assert.True(map.Optional);
        Assert.Same(map, map.WithReaderFactory(factory));
        Assert.True(map.Optional);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void WithReaderFactory_NullReader_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("readerFactory", () => map.WithReaderFactory(null!));
    }

    [Fact]
    public void WithCellValueMappers_ValidMappers_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);
        ICellMapper mapper1 = Assert.Single(map.Pipeline.CellValueMappers);
        ICellMapper mapper2 = new BoolMapper();

        Assert.Same(map, map.WithCellValueMappers(mapper2));
        Assert.Equal([mapper1, mapper2], map.Pipeline.CellValueMappers);
    }

    [Fact]
    public void WithCellValueMappers_NullMappers_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("mappers", () => map.WithCellValueMappers(null!));
    }

    [Fact]
    public void WithCellValueMappers_NullMapperInMapperss_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("mappers", () => map.WithCellValueMappers([null!]));
    }

    [Fact]
    public void WithMapping_ValidReader_Success()
    {
        var dictionaryMapping = new Dictionary<string, DateTime> { { "key", DateTime.MinValue } };
        StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;

        OneToOneMap<DateTime> map = Map(t => t.DateValue);
        Assert.Same(map, map.WithMapping(dictionaryMapping, comparer));

        DictionaryMapper<DateTime> item = map.Pipeline.CellValueMappers.OfType<DictionaryMapper<DateTime>>().Single();
        Assert.NotSame(dictionaryMapping, item.MappingDictionary);
        Assert.Equal(dictionaryMapping, item.MappingDictionary);

        Assert.Same(comparer, Assert.IsType<Dictionary<string, DateTime>>(item.MappingDictionary).Comparer);
    }

    [Fact]
    public void WithMapping_NullMapping_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("mappingDictionary", () => map.WithMapping((Dictionary<string, string>)null!));
    }

    [Fact]
    public void MakeOptional_HasMapper_ReturnsExpected()
    {
        var factory = new ColumnIndexReaderFactory(1);
        var map = Map(t => t.Value).WithReaderFactory(factory);
        Assert.False(map.Optional);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void MakeOptional_AlreadyOptional_ReturnsExpected()
    {
        var factory = new ColumnIndexReaderFactory(1);
        var map = Map(t => t.Value).WithReaderFactory(factory);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void MakePreserveFormatting_HasMapper_ReturnsExpected()
    {
        var factory = new ColumnIndexReaderFactory(1);
        var map = Map(t => t.Value).WithReaderFactory(factory);
        Assert.False(map.PreserveFormatting);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void MakePreserveFormatting_AlreadyPreserveFormatting_ReturnsExpected()
    {
        var factory = new ColumnIndexReaderFactory(1);
        var map = Map(t => t.Value).WithReaderFactory(factory);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void WithTrim_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithTrim());

        ICellTransformer transformer = Assert.Single(map.Pipeline.CellValueTransformers);
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

        OneToOneMap<DateTime> map = Map(t => t.DateValue);
        Assert.Same(map, map.WithDateFormats(formatsArray));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        OneToOneMap<DateTime> map = Map(t => t.DateValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithDateFormats(formatsArray));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_AutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        OneToOneMap<DateTime> map = Map(t => t.DateValue);
        Assert.Same(map, map.WithDateFormats(formats));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        OneToOneMap<DateTime> map = Map(t => t.DateValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithDateFormats(formats));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithDateFormats_NullFormats_ThrowsArgumentNullException()
    {
        OneToOneMap<DateTime> map = Map(t => t.DateValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithDateFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithDateFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithDateFormats_EmptyFormats_ThrowsArgumentException()
    {
        OneToOneMap<DateTime> map = Map(t => t.DateValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats(new List<string>()));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NullableAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);
        Assert.Same(map, map.WithDateFormats(formatsArray));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NullableNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithDateFormats(formatsArray));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NullableAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);
        Assert.Same(map, map.WithDateFormats(formats));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithDateFormats_NullableNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithDateFormats(formats));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
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
        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats(new List<string>()));
    }

    [Fact]
    public void WithConverter_SuccessConverterSimple_ReturnsExpected()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithConverter(stringValue =>
        {
            Assert.Equal("stringValue", stringValue);
            return "abc";
        }));
        ConvertUsingMapper item = map.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

        CellMapperResult result = item.Converter(new ReadCellResult(-1, "stringValue", preserveFormatting: false));
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

        var map = Map(t => t.Value);
        Assert.Same(map, map.WithConverter(converter));
        ConvertUsingMapper item = map.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

        CellMapperResult result = item.Converter(new ReadCellResult(-1, "stringValue", preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.IsType<NotSupportedException>(result.Exception);
    }

    [Fact]
    public void WithConverter_SuccessConverterAdvanced_ReturnsExpected()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithConverter(readResult =>
        {
            Assert.Equal("stringValue", readResult.StringValue);
            return CellMapperResult.Success("abc");
        }));
        ConvertUsingMapper item = map.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

        CellMapperResult result = item.Converter(new ReadCellResult(-1, "stringValue", preserveFormatting: false));
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

        var map = Map(t => t.Value);
        Assert.Same(map, map.WithConverter(converter));
        ConvertUsingMapper item = map.Pipeline.CellValueMappers.OfType<ConvertUsingMapper>().Single();

        Assert.Throws<NotSupportedException>(() => item.Converter(new ReadCellResult(-1, "stringValue", preserveFormatting: false)));
    }

    [Fact]
    public void WithConverter_NullConverter_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter((ConvertUsingSimpleMapperDelegate<string>)null!));
        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter((ConvertUsingMapperDelegate)null!));
    }

    [Fact]
    public void WithValueFallback_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithValueFallback("abc"));

        FixedValueFallback emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        FixedValueFallback invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);

        Assert.Equal("abc", emptyFallback.Value);
        Assert.Equal("abc", invalidFallback.Value);
    }

    [Fact]
    public void WithThrowingFallback_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithThrowingFallback());

        Assert.IsType<ThrowFallback>(map.Pipeline.EmptyFallback);
        Assert.IsType<ThrowFallback>(map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithEmptyFallback_Invoke_Success()
    {
        var mapping = Map(t => t.Value);
        Assert.Same(mapping, mapping.WithEmptyFallback("abc"));

        FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(mapping.Pipeline.EmptyFallback);
        Assert.Equal("abc", fallback.Value);
    }

    [Fact]
    public void WithThrowingEmptyFallback_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithThrowingEmptyFallback());

        Assert.IsType<ThrowFallback>(map.Pipeline.EmptyFallback);
    }

    [Fact]
    public void WithEmptyFallbackItem_ValidFallbackItem_Success()
    {
        IFallbackItem fallback = new FixedValueFallback(10);
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithEmptyFallbackItem(fallback));

        Assert.Same(fallback, map.Pipeline.EmptyFallback);
    }

    [Fact]
    public void WithEmptyFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);

        Assert.Throws<ArgumentNullException>("fallbackItem", () => map.WithEmptyFallbackItem(null!));
    }

    [Fact]
    public void WithInvalidFallback_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithInvalidFallback("abc"));

        FixedValueFallback fallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal("abc", fallback.Value);
    }

    [Fact]
    public void WithThrowingInvalidFallback_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithThrowingInvalidFallback());

        Assert.IsType<ThrowFallback>(map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithInvalidFallbackItem_ValidFallbackItem_Success()
    {
        IFallbackItem fallback = new FixedValueFallback(10);
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithInvalidFallbackItem(fallback));

        Assert.Same(fallback, map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithInvalidFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("fallbackItem", () => map.WithInvalidFallbackItem(null!));
    }
}
