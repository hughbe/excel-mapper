using System;
using System.Collections.Generic;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnIndicesReaderFactoryTests
{
    [Fact]
    public void Ctor_ColumnIndices()
    {
        var columnIndices = new int[] { 0, 1 };
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

    public static IEnumerable<object[]> GetReader_TestData()
    {
        yield return new object[] { new int[] { 0 } };
        yield return new object[] { new int[] { 0, 0 } };
    }

    [Theory]
    [MemberData(nameof(GetReader_TestData))]
    public void GetReader_InvokeSheetWithHeading_ReturnsExpected(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetReader(sheet));
        Assert.Equal(columnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetReader(sheet));
    }

    public static IEnumerable<object[]> GetReader_NoSuchColumn_TestData()
    {
        yield return new object[] { new int[] { 1 } };
        yield return new object[] { new int[] { 0, 1 } };
    }

    [Theory]
    [MemberData(nameof(GetReader_NoSuchColumn_TestData))]
    public void GetReader_InvokeNoMatch_ReturnsNull(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        Assert.Null(factory.GetReader(sheet));
    }

    [Theory]
    [MemberData(nameof(GetReader_TestData))]
    public void GetReader_InvokeSheetWithNoHeadingHasHeading_ReturnsExpected(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetReader(sheet));
        Assert.Equal(columnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Theory]
    [MemberData(nameof(GetReader_TestData))]
    public void GetReader_InvokeSheetWithNoHeadingHasNoHeading_ReturnsExpected(int[] columnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnIndicesReaderFactory(columnIndices);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetReader(sheet));
        Assert.Equal(columnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnIndicesReaderFactory(0);
        Assert.Throws<ArgumentNullException>(() => factory.GetReader(null!));
    }
}
