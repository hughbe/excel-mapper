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
        var mapper1 = Assert.Single(map.Pipeline.CellValueMappers);
        var mapper2 = new BoolMapper();

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

        var map = Map(t => t.DateTimeValue);
        Assert.Same(map, map.WithMapping(dictionaryMapping, comparer));

        var item = map.Pipeline.CellValueMappers.OfType<DictionaryMapper<DateTime>>().Single();
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

        var transformer = Assert.Single(map.Pipeline.CellValueTransformers);
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
    public void WithFormats_DateTimeAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.DateTimeValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.DateTimeValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.DateTimeValue);
        Assert.Same(map, map.WithFormats(formats));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.DateTimeValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_DateTimeNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.DateTimeValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateTimeEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateTimeNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateTimeEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeNullableAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableDateTimeValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeNullableNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableDateTimeValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeNullableAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableDateTimeValue);
        Assert.Same(map, map.WithFormats(formats));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeNullableNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableDateTimeValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        DateTimeMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_DateTimeNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableDateTimeValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateTimeNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateTimeNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateTimeNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeOffsetAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.DateTimeOffsetValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        DateTimeOffsetMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeOffsetMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeOffsetNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.DateTimeOffsetValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        DateTimeOffsetMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeOffsetMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeOffsetAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.DateTimeOffsetValue);
        Assert.Same(map, map.WithFormats(formats));

        DateTimeOffsetMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeOffsetMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeOffsetNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.DateTimeOffsetValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        DateTimeOffsetMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeOffsetMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.DateTimeOffsetValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeOffsetValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeOffsetValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeOffsetValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeOffsetNullableAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableDateTimeOffsetValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        DateTimeOffsetMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeOffsetMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeOffsetNullableNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableDateTimeOffsetValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        DateTimeOffsetMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeOffsetMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeOffsetNullableAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);
        Assert.Same(map, map.WithFormats(formats));

        DateTimeOffsetMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeOffsetMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateTimeOffsetNullableNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        DateTimeOffsetMapper item = map.Pipeline.CellValueMappers.OfType<DateTimeOffsetMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeSpanAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.TimeSpanValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        TimeSpanMapper item = map.Pipeline.CellValueMappers.OfType<TimeSpanMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeSpanNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.TimeSpanValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        TimeSpanMapper item = map.Pipeline.CellValueMappers.OfType<TimeSpanMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeSpanAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.TimeSpanValue);
        Assert.Same(map, map.WithFormats(formats));

        TimeSpanMapper item = map.Pipeline.CellValueMappers.OfType<TimeSpanMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeSpanNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.TimeSpanValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        TimeSpanMapper item = map.Pipeline.CellValueMappers.OfType<TimeSpanMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_TimeSpanNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.TimeSpanValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_TimeSpanEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeSpanValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_TimeSpanNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeSpanValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_TimeSpanEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeSpanValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeSpanNullableAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableTimeSpanValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        TimeSpanMapper item = map.Pipeline.CellValueMappers.OfType<TimeSpanMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeSpanNullableNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableTimeSpanValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        TimeSpanMapper item = map.Pipeline.CellValueMappers.OfType<TimeSpanMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeSpanNullableAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableTimeSpanValue);
        Assert.Same(map, map.WithFormats(formats));

        TimeSpanMapper item = map.Pipeline.CellValueMappers.OfType<TimeSpanMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeSpanNullableNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableTimeSpanValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        TimeSpanMapper item = map.Pipeline.CellValueMappers.OfType<TimeSpanMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_TimeSpanNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableTimeSpanValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_TimeSpanNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeSpanValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_TimeSpanNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeSpanValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_TimeSpanNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeSpanValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateOnlyAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.DateOnlyValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        DateOnlyMapper item = map.Pipeline.CellValueMappers.OfType<DateOnlyMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateOnlyNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.DateOnlyValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        DateOnlyMapper item = map.Pipeline.CellValueMappers.OfType<DateOnlyMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateOnlyAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.DateOnlyValue);
        Assert.Same(map, map.WithFormats(formats));

        DateOnlyMapper item = map.Pipeline.CellValueMappers.OfType<DateOnlyMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateOnlyNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.DateOnlyValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        DateOnlyMapper item = map.Pipeline.CellValueMappers.OfType<DateOnlyMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_DateOnlyNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.DateOnlyValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateOnlyEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateOnlyNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateOnlyEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateOnlyNullableAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableDateOnlyValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        DateOnlyMapper item = map.Pipeline.CellValueMappers.OfType<DateOnlyMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateOnlyNullableNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableDateOnlyValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        DateOnlyMapper item = map.Pipeline.CellValueMappers.OfType<DateOnlyMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateOnlyNullableAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableDateOnlyValue);
        Assert.Same(map, map.WithFormats(formats));

        DateOnlyMapper item = map.Pipeline.CellValueMappers.OfType<DateOnlyMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_DateOnlyNullableNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableDateOnlyValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        DateOnlyMapper item = map.Pipeline.CellValueMappers.OfType<DateOnlyMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_DateOnlyNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableDateOnlyValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateOnlyNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateOnlyNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateOnlyNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeOnlyAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.TimeOnlyValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        TimeOnlyMapper item = map.Pipeline.CellValueMappers.OfType<TimeOnlyMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeOnlyNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.TimeOnlyValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        TimeOnlyMapper item = map.Pipeline.CellValueMappers.OfType<TimeOnlyMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeOnlyAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.TimeOnlyValue);
        Assert.Same(map, map.WithFormats(formats));

        TimeOnlyMapper item = map.Pipeline.CellValueMappers.OfType<TimeOnlyMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeOnlyNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.TimeOnlyValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        TimeOnlyMapper item = map.Pipeline.CellValueMappers.OfType<TimeOnlyMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_TimeOnlyNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.TimeOnlyValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_TimeOnlyEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_TimeOnlyNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_TimeOnlyEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeOnlyNullableAutoMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableTimeOnlyValue);
        Assert.Same(map, map.WithFormats(formatsArray));

        TimeOnlyMapper item = map.Pipeline.CellValueMappers.OfType<TimeOnlyMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeOnlyNullableNotMappedStringArray_Success(IEnumerable<string> formats)
    {
        var formatsArray = formats.ToArray();

        var map = Map(t => t.NullableTimeOnlyValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formatsArray));

        TimeOnlyMapper item = map.Pipeline.CellValueMappers.OfType<TimeOnlyMapper>().Single();
        Assert.Same(formatsArray, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeOnlyNullableAutoMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableTimeOnlyValue);
        Assert.Same(map, map.WithFormats(formats));

        TimeOnlyMapper item = map.Pipeline.CellValueMappers.OfType<TimeOnlyMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Theory]
    [MemberData(nameof(Formats_TestData))]
    public void WithFormats_TimeOnlyNullableNotMappedIEnumerableString_Success(ICollection<string> formats)
    {
        var map = Map(t => t.NullableTimeOnlyValue);
        map.Pipeline.RemoveCellValueMapper(0);

        Assert.Same(map, map.WithFormats(formats));

        TimeOnlyMapper item = map.Pipeline.CellValueMappers.OfType<TimeOnlyMapper>().Single();
        Assert.Equal(formats, item.Formats);
    }

    [Fact]
    public void WithFormats_TimeOnlyNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableTimeOnlyValue);

        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_TimeOnlyNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_TimeOnlyNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_TimeOnlyNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeOnlyValue);

        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
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

        var result = item.Converter(new ReadCellResult(0, "stringValue", preserveFormatting: false));
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

        var result = item.Converter(new ReadCellResult(0, "stringValue", preserveFormatting: false));
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

        var result = item.Converter(new ReadCellResult(0, "stringValue", preserveFormatting: false));
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

        Assert.Throws<NotSupportedException>(() => item.Converter(new ReadCellResult(0, "stringValue", preserveFormatting: false)));
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
