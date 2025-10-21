using System;
using System.Collections.Generic;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnIndicesReaderFactoryTests
{
    public static IEnumerable<object[]> Ctor_ParamsInt_TestData()
    {
        yield return new object[] { new int[] { 0 } };
        yield return new object[] { new int[] { 0, 0 } };
        yield return new object[] { new int[] { 0, 1 } };
    }

    [Theory]
    [MemberData(nameof(Ctor_ParamsInt_TestData))]
    public void Ctor_ParamsInt(int[] columnIndices)
    {
        var reader = new ColumnIndicesReaderFactory(columnIndices);
        Assert.Same(columnIndices, reader.ColumnIndices);
    }

    [Fact]
    public void Ctor_NullColumnIndices_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnIndices", () => new ColumnIndicesReaderFactory(null!));
    }

    [Fact]
    public void Ctor_EmptyColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnIndices", () => new ColumnIndicesReaderFactory([]));
    }

    [Fact]
    public void Ctor_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => new ColumnIndicesReaderFactory([-1]));
    }    

    public static IEnumerable<object[]> GetCellReader_TestData()
    {
        yield return new object[] { new int[] { 0 }, 0 };
        yield return new object[] { new int[] { 0, 0 }, 0 };
        yield return new object[] { new int[] { 1, 0 }, 0 };
        yield return new object[] { new int[] { int.MaxValue, 0 }, 0 };
    }

    [Theory]
    [MemberData(nameof(GetCellReader_TestData))]
    public void GetCellReader_InvokeSheetWithHeading_ReturnsExpected(int[] columnIndices, int expectedColumnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(expectedColumnIndex, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
    }

    public static IEnumerable<object[]> GetCellReader_NoSuchColumn_TestData()
    {
        yield return new object[] { new int[] { 1 } };
        yield return new object[] { new int[] { int.MaxValue } };
    }

    [Theory]
    [MemberData(nameof(GetCellReader_NoSuchColumn_TestData))]
    public void GetCellReader_InvokeNoMatch_ReturnsNull(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        Assert.Null(factory.GetCellReader(sheet));
    }

    [Theory]
    [MemberData(nameof(GetCellReader_TestData))]
    public void GetCellReader_InvokeSheetWithNoHeadingHasHeading_ReturnsExpected(int[] columnIndices, int expectedColumnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(expectedColumnIndex, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Theory]
    [MemberData(nameof(GetCellReader_TestData))]
    public void GetCellReader_InvokeSheetWithNoHeadingHasNoHeading_ReturnsExpected(int[] columnIndices, int expectedColumnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(expectedColumnIndex, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnIndicesReaderFactory(0);
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellReader(null!));
    }

    public static IEnumerable<object[]> GetCellsReader_TestData()
    {
        yield return new object[] { new int[] { 0 } };
        yield return new object[] { new int[] { 0, 0 } };
    }

    [Theory]
    [MemberData(nameof(GetCellsReader_TestData))]
    public void GetCellsReader_InvokeSheetWithHeading_ReturnsExpected(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetCellsReader(sheet));
        Assert.Equal(columnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetCellsReader(sheet));
    }

    public static IEnumerable<object[]> GetCellsReader_NoSuchColumn_TestData()
    {
        yield return new object[] { new int[] { 1 } };
        yield return new object[] { new int[] { 0, 1 } };
    }

    [Theory]
    [MemberData(nameof(GetCellsReader_NoSuchColumn_TestData))]
    public void GetCellsReader_InvokeNoMatch_ReturnsNull(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        Assert.Null(factory.GetCellsReader(sheet));
    }

    [Theory]
    [MemberData(nameof(GetCellsReader_TestData))]
    public void GetCellsReader_InvokeSheetWithNoHeadingHasHeading_ReturnsExpected(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetCellsReader(sheet));
        Assert.Equal(columnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetCellsReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Theory]
    [MemberData(nameof(GetCellsReader_TestData))]
    public void GetCellsReader_InvokeSheetWithNoHeadingHasNoHeading_ReturnsExpected(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetCellsReader(sheet));
        Assert.Equal(columnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetCellsReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellsReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnIndicesReaderFactory(0);
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellsReader(null!));
    }

#pragma warning disable CS0184 // The is operator is being used to test interface implementation
    [Fact]
    public void Interfaces_IColumnNameProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new ColumnIndicesReaderFactory(0);
        Assert.False(factory is IColumnNameProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndexProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new ColumnIndicesReaderFactory(0);
        Assert.False(factory is IColumnIndexProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnNamesProviderCellReaderFactory_DoesImplement()
    {
        var factory = new ColumnIndicesReaderFactory(0);
        Assert.False(factory is IColumnNamesProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndicesProviderCellReaderFactory_DoesImplement()
    {
        var factory = new ColumnIndicesReaderFactory(0);
        Assert.True(factory is IColumnIndicesProviderCellReaderFactory);
    }
#pragma warning restore CS0184

    [Fact]
    public void GetColumnIndices_Invoke_ReturnsExpected()
    {
        var factory = new ColumnIndicesReaderFactory([1, 2, 3]);
        Assert.Equal([1, 2, 3], factory.GetColumnIndices(null!)!);
    }
}
