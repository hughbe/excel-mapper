using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests;

public class OneToOneMapTTests
{
    [Fact]
    public void Ctor_ICellReaderFactory()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<int>(factory);
        Assert.Same(factory, map.ReaderFactory);
        Assert.Empty(map.CellValueMappers);
        Assert.Same(map.CellValueMappers, map.CellValueMappers);
        Assert.Same(map.CellValueMappers, map.Pipeline.CellValueMappers);
        Assert.Empty(map.CellValueTransformers);
        Assert.Same(map.CellValueTransformers, map.CellValueTransformers);
        Assert.Same(map.CellValueTransformers, map.Pipeline.CellValueTransformers);
        Assert.False(map.Optional);
        Assert.False(map.PreserveFormatting);
    }

    [Fact]
    public void Ctor_NullReader_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("readerFactory", () => new SubOneToOneMap<int>(null!));
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
        var map = new SubOneToOneMap<int>(factory)
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
        var map = new SubOneToOneMap<int>(factory);

        Assert.Throws<ArgumentNullException>("value", () => map.ReaderFactory = null!);
    }

    [Fact]
    public void EmptyFallback_Set_GetReturnsExpected()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<int>(factory);

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
        var map = new SubOneToOneMap<int>(factory);

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
        var map = new SubOneToOneMap<int>(factory)
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
        var map = new SubOneToOneMap<int>(factory)
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

    [Fact]
    public void AddCellValueMapper_ValidItem_Success()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<int>(factory);
        var item1 = new BoolMapper();
        var item2 = new BoolMapper();

        map.AddCellValueMapper(item1);
        map.AddCellValueMapper(item2);
        Assert.Equal([item1, item2], map.CellValueMappers);
    }

    [Fact]
    public void AddCellValueMapper_NullItem_ThrowsArgumentNullException()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<int>(factory);

        Assert.Throws<ArgumentNullException>("mapper", () => map.AddCellValueMapper(null!));
    }

    [Fact]
    public void RemoveCellValueMapper_Index_Success()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<int>(factory);
        map.AddCellValueMapper(new BoolMapper());

        map.RemoveCellValueMapper(0);
        Assert.Empty(map.CellValueMappers);
    }

    [Fact]
    public void AddCellValueTransformer_ValidTransformer_Success()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<int>(factory);
        var transformer1 = new TrimCellValueTransformer();
        var transformer2 = new TrimCellValueTransformer();

        map.AddCellValueTransformer(transformer1);
        map.AddCellValueTransformer(transformer2);
        Assert.Equal([transformer1, transformer2], map.CellValueTransformers);
    }

    [Fact]
    public void AddCellValueTransformer_NullTransformer_ThrowsArgumentNullException()
    {
        var factory = new ColumnNameReaderFactory("Column");
        var map = new SubOneToOneMap<int>(factory);
        Assert.Throws<ArgumentNullException>("transformer", () => map.AddCellValueTransformer(null!));
    }

    [Fact]
    public void TryGetValue_InvokeCantReadPropertyInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockSingleCellValueReader(() => (false, default)));
        var map = new SubOneToOneMap<int>(factory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadFieldInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockSingleCellValueReader(() => (false, default)));
        var map = new SubOneToOneMap<int>(factory);
        MemberInfo member = typeof(TestClass).GetField(nameof(TestClass._field))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadEventInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockSingleCellValueReader(() => (false, default)));
        var map = new SubOneToOneMap<int>(factory);
        MemberInfo member = typeof(TestClass).GetEvent(nameof(TestClass.Event))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMember_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockSingleCellValueReader(() => (false, default)));
        var map = new SubOneToOneMap<int>(factory);
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
