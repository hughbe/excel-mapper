using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class OneToOneMapTTests
{
    [Fact]
    public void Ctor_ICellReaderFactory()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory);
        Assert.Same(factory, map.ReaderFactory);
        Assert.Empty(map.Mappers);
        Assert.Same(map.Mappers, map.Mappers);
        Assert.Same(map.Mappers, map.Pipeline.Mappers);
        Assert.Empty(map.Transformers);
        Assert.Same(map.Transformers, map.Transformers);
        Assert.Same(map.Transformers, map.Pipeline.Transformers);
        Assert.False(map.Optional);
        Assert.False(map.PreserveFormatting);
    }

    [Fact]
    public void Ctor_NullReader_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("readerFactory", () => new SubOneToOneMap<string>(null!));
    }

    public static IEnumerable<object[]> ReaderFactory_Set_TestData()
    {
        yield return new object[] { new ColumnNameReaderFactory("Column") };
    }

    [Theory]
    [MemberData(nameof(ReaderFactory_Set_TestData))]
    public void ReaderFactory_SetValid_GetReturnsExpected(ICellReaderFactory value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory)
        {
            ReaderFactory = value
        };
        Assert.Same(value, map.ReaderFactory);

        // Set same.
        map.ReaderFactory = value;
        Assert.Same(value, map.ReaderFactory);
    }

    [Fact]
    public void ReaderFactory_SetNull_ThrowsArgumentNullException()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory);

        Assert.Throws<ArgumentNullException>("value", () => map.ReaderFactory = null!);
    }

    [Fact]
    public void EmptyFallback_Set_GetReturnsExpected()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory);

        var fallback = new FixedValueFallback(10);
        map.EmptyFallback = fallback;
        Assert.Same(fallback, map.EmptyFallback);

        map.EmptyFallback = null;
        Assert.Null(map.EmptyFallback);
    }

    [Fact]
    public void InvalidFallback_Set_GetReturnsExpected()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory);

        var fallback = new FixedValueFallback(10);
        map.InvalidFallback = fallback;
        Assert.Same(fallback, map.InvalidFallback);

        map.InvalidFallback = null;
        Assert.Null(map.InvalidFallback);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Optional_Set_GetReturnsExpected(bool value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory)
        {
            Optional = value
        };
        Assert.Equal(value, map.Optional);

        // Set same.
        map.Optional = value;
        Assert.Equal(value, map.Optional);

        // Set different.
        map.Optional = !value;
        Assert.Equal(!value, map.Optional);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PreserveFormatting_Set_GetReturnsExpected(bool value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory)
        {
            PreserveFormatting = value
        };
        Assert.Equal(value, map.PreserveFormatting);

        // Set same.
        map.PreserveFormatting = value;
        Assert.Equal(value, map.PreserveFormatting);

        // Set different.
        map.PreserveFormatting = !value;
        Assert.Equal(!value, map.PreserveFormatting);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithValueFallback_Invoke_Success(string? value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory);
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
    public void WithValueFallback_InvokeWithFallbacks_Success(string? value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory).WithEmptyFallback("Empty").WithInvalidFallback("Invalid");
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
    public void WithEmptyFallback_Invoke_Success(string? value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory);
        Assert.Same(map, map.WithEmptyFallback(value));

        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        Assert.Equal(value, emptyFallback.Value);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithEmptyFallback_InvokeWithInvalidFallback_Success(string? value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory).WithInvalidFallback("Invalid");
        Assert.Same(map, map.WithEmptyFallback(value));

        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        Assert.Equal(value, emptyFallback.Value);
        
        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal("Invalid", invalidFallback.Value);
    }
    

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithInvalidFallback_Invoke_Success(string? value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory);
        Assert.Same(map, map.WithInvalidFallback(value));

        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal(value, invalidFallback.Value);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData(null)]
    public void WithInvalidFallback_InvokeWithEmptyFallback_Success(string? value)
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<string>(factory).WithEmptyFallback("Empty");
        Assert.Same(map, map.WithInvalidFallback(value));

        var invalidFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.InvalidFallback);
        Assert.Equal(value, invalidFallback.Value);

        var emptyFallback = Assert.IsType<FixedValueFallback>(map.Pipeline.EmptyFallback);
        Assert.Equal("Empty", emptyFallback.Value);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadPropertyInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockSingleCellValueReader(() => (false, default)));
        var map = new SubOneToOneMap<string>(factory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadFieldInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockSingleCellValueReader(() => (false, default)));
        var map = new SubOneToOneMap<string>(factory);
        MemberInfo member = typeof(TestClass).GetField(nameof(TestClass._field))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadEventInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockSingleCellValueReader(() => (false, default)));
        var map = new SubOneToOneMap<string>(factory);
        MemberInfo member = typeof(TestClass).GetEvent(nameof(TestClass.Event))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMember_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockSingleCellValueReader(() => (false, default)));
        var map = new SubOneToOneMap<string>(factory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnNameProviderCellReaderFactory_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNameReaderFactory("NoSuchColumn");
        var map = new SubOneToOneMap<string>(factory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnIndexProviderCellReaderFactory_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Reader.Read();

        var factory = new ColumnIndexReaderFactory(int.MaxValue);
        var map = new SubOneToOneMap<string>(factory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    private class SubOneToOneMap<T> : OneToOneMap<T>
    {
        public SubOneToOneMap(ICellReaderFactory factory) : base(factory)
        {
        }
    }

    private class MockSingleCellValueReader(Func<(bool, ReadCellResult)> action) : ICellReader
    {
        public Func<(bool, ReadCellResult)> Action { get; } = action;

        public bool TryGetValue(IExcelDataReader factory, bool preserveFormatting, out ReadCellResult result)
        {   
            var (ret, res) = Action();
            result = res;
            return ret;
        }
    }

    private class MockReaderFactory(ICellReader Result) : ICellReaderFactory
    {
        public ICellReader? GetCellReader(ExcelSheet sheet) => Result;
    }

    private class TestClass
    {
        public string Value { get; set; } = default!;
#pragma warning disable 0649
        public string _field = default!;
#pragma warning restore 0649

        public event EventHandler Event { add { } remove { } }
    }
}
