using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests;

public class ValuePipelineExtensionsTests : ExcelClassMap<Helpers.TestClass>
{
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

    [Fact]
    public void WithDateFormats_NullValueInFormats_ThrowsArgumentException()
    {
        OneToOneMap<DateTime> map = Map(t => t.DateValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithDateFormats_EmptyValueInFormats_ThrowsArgumentException()
    {
        OneToOneMap<DateTime> map = Map(t => t.DateValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats(new List<string> { string.Empty }));
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
        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithDateFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithDateFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithDateFormats_NullableEmptyFormats_ThrowsArgumentException()
    {
        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats(new List<string>()));
    }

    [Fact]
    public void WithDateFormats_NullableNullValueInFormats_ThrowsArgumentException()
    {
        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithDateFormats_NullableEmptyValueInFormats_ThrowsArgumentException()
    {
        OneToOneMap<DateTime?> map = Map(t => t.NullableDateValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithDateFormats(new List<string> { string.Empty }));
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
        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter(null!));
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithValueFallback_Invoke_Success(object? value)
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithValueFallback(value));

        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);

        Assert.Same(emptyFallback, invalidFallback);
        Assert.Equal(value, emptyFallback.Value);
        Assert.Equal(value, invalidFallback.Value);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithValueFallback_InvokeWithFallbacks_Success(object? value)
    {
        var map = Map(t => t.Value).WithEmptyFallback("Empty").WithInvalidFallback("Invalid");
        Assert.Same(map, map.WithValueFallback(value));

        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);

        Assert.Same(emptyFallback, invalidFallback);
        Assert.Equal(value, emptyFallback.Value);
        Assert.Equal(value, invalidFallback.Value);
    }

    [Fact]
    public void WithThrowingFallback_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithThrowingFallback());

        Assert.Same(map.Pipeline.EmptyFallback, map.Pipeline.InvalidFallback);
        Assert.IsType<ThrowFallback>(map.Pipeline.EmptyFallback);
        Assert.IsType<ThrowFallback>(map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithThrowingFallback_InvokeWithFallbacks_Success()
    {
        var map = Map(t => t.Value).WithEmptyFallback("Empty").WithInvalidFallback("Invalid");
        Assert.Same(map, map.WithThrowingFallback());

        Assert.Same(map.Pipeline.EmptyFallback, map.Pipeline.InvalidFallback);
        Assert.IsType<ThrowFallback>(map.Pipeline.EmptyFallback);
        Assert.IsType<ThrowFallback>(map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithFallbackItem_Invoke_Success()
    {
        var fallback = new FixedValueFallback("Fallback");
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithFallbackItem(fallback));

        Assert.Same(map.Pipeline.EmptyFallback, map.Pipeline.InvalidFallback);
        Assert.Same(fallback, map.Pipeline.EmptyFallback);
        Assert.Same(fallback, map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithFallbackItem_InvokeWithFallbacks_Success()
    {
        var fallback = new FixedValueFallback("Value");
        var map = Map(t => t.Value).WithEmptyFallback("Empty").WithInvalidFallback("Invalid");
        Assert.Same(map, map.WithFallbackItem(fallback));

        Assert.Same(map.Pipeline.EmptyFallback, map.Pipeline.InvalidFallback);
        Assert.Same(fallback, map.Pipeline.EmptyFallback);
        Assert.Same(fallback, map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("fallbackItem", () => map.WithFallbackItem(null!));
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithEmptyFallback_Invoke_Success(object? value)
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithEmptyFallback(value));

        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        Assert.Equal(value, emptyFallback.Value);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithEmptyFallback_InvokeWithInvalidFallback_Success(object? value)
    {
        var map = Map(t => t.Value).WithInvalidFallback("Invalid");
        Assert.Same(map, map.WithEmptyFallback(value));

        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        Assert.Equal(value, emptyFallback.Value);
        
        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal("Invalid", invalidFallback.Value);
    }

    [Fact]
    public void WithThrowingEmptyFallback_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithThrowingEmptyFallback());

        Assert.IsType<ThrowFallback>(map.Pipeline.EmptyFallback);
    }

    [Fact]
    public void WithThrowingEmptyFallback_InvokeWithInvalidFallback_Success()
    {
        var map = Map(t => t.Value).WithInvalidFallback("Invalid");
        Assert.Same(map, map.WithThrowingEmptyFallback());

        Assert.IsType<ThrowFallback>(map.Pipeline.EmptyFallback);

        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal("Invalid", invalidFallback.Value);
    }

    [Fact]
    public void WithEmptyFallbackItem_Invoke_Success()
    {
        var emptyFallback = new FixedValueFallback(10);
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithEmptyFallbackItem(emptyFallback));

        Assert.Same(emptyFallback, map.Pipeline.EmptyFallback);
    }

    [Fact]
    public void WithEmptyFallbackItem_InvokeWithInvalidFallbackItem_Success()
    {
        var emptyFallback = new FixedValueFallback(10);
        var map = Map(t => t.Value).WithInvalidFallback("Invalid");
        Assert.Same(map, map.WithEmptyFallbackItem(emptyFallback));

        Assert.Same(emptyFallback, map.Pipeline.EmptyFallback);

        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal("Invalid", invalidFallback.Value);        
    }

    [Fact]
    public void WithEmptyFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("fallbackItem", () => map.WithEmptyFallbackItem(null!));
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithInvalidFallback_Invoke_Success(object? value)
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithInvalidFallback(value));

        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal(value, invalidFallback.Value);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithInvalidFallback_InvokeWithEmptyFallback_Success(object? value)
    {
        var map = Map(t => t.Value).WithEmptyFallback("Empty");
        Assert.Same(map, map.WithInvalidFallback(value));

        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal(value, invalidFallback.Value);

        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        Assert.Equal("Empty", emptyFallback.Value);
    }

    [Fact]
    public void WithThrowingInvalidFallback_Invoke_Success()
    {
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithThrowingInvalidFallback());

        Assert.IsType<ThrowFallback>(map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithThrowingInvalidFallback_InvokeWithEmptyFallback_Success()
    {
        var map = Map(t => t.Value).WithEmptyFallback("Empty");
        Assert.Same(map, map.WithThrowingInvalidFallback());

        Assert.IsType<ThrowFallback>(map.Pipeline.InvalidFallback);
        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        Assert.Equal("Empty", emptyFallback.Value);
    }

    [Fact]
    public void WithInvalidFallbackItem_Invoke_Success()
    {
        var invalidFallback = new FixedValueFallback(10);
        var map = Map(t => t.Value);
        Assert.Same(map, map.WithInvalidFallbackItem(invalidFallback));

        Assert.Same(invalidFallback, map.Pipeline.InvalidFallback);
    }

    [Fact]
    public void WithInvalidFallbackItem_InvokeWithEmptyFallback_Success()
    {
        var invalidFallback = new FixedValueFallback(10);
        var map = Map(t => t.Value).WithEmptyFallback("Empty");
        Assert.Same(map, map.WithInvalidFallbackItem(invalidFallback));

        Assert.Same(invalidFallback, map.Pipeline.InvalidFallback);
        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        Assert.Equal("Empty", emptyFallback.Value);
    }

    [Fact]
    public void WithInvalidFallbackItem_NullFallbackItem_ThrowsArgumentNullException()
    {
        var map = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("fallbackItem", () => map.WithInvalidFallbackItem(null!));
    }
}
