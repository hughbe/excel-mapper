using System;
using System.Linq;
using Xunit;

namespace ExcelMapper.Tests;

public class ExcelHeadingTests
{
    [Fact]
    public void GetColumnName_Invoke_Roundtrips()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["Int Value", "StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "MappedValue", "TrimmedValue"], heading.ColumnNames);

        string[] columnNames = [.. heading.ColumnNames];
        for (int i = 0; i < columnNames.Length; i++)
        {
            Assert.Equal(i, heading.GetColumnIndex(columnNames[i]));
            Assert.Equal(columnNames[i], heading.GetColumnName(i));
        }
    }

    [Fact]
    public void GetColumnName_DuplicatedColumn_Success()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["MyColumn", "MyColumn"], heading.ColumnNames);
        Assert.Equal("MyColumn", heading.GetColumnName(0));
        Assert.Equal("MyColumn", heading.GetColumnName(1));
    }

    [Fact]
    public void GetColumnName_DuplicatedColumnEmpty_Success()
    {
        using var importer = Helpers.GetImporter("EmptyColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["", "Column2", "", " Column4 "], heading.ColumnNames);
        Assert.Equal("", heading.GetColumnName(0));
        Assert.Equal("Column2", heading.GetColumnName(1));
        Assert.Equal("", heading.GetColumnName(2));
        Assert.Equal(" Column4 ", heading.GetColumnName(3));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(8)]
    public void GetColumnName_InvalidIndex_ThrowsArgumentOutOfRangeException(int columnIndex)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => heading.GetColumnName(columnIndex));
    }

    [Fact]
    public void GetColumnIndex_InvokeValidColumnName_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Equal(1, heading.GetColumnIndex("StringValue"));
        Assert.Equal(1, heading.GetColumnIndex("stringvalue"));
    }

    [Fact]
    public void GetColumnIndex_DuplicatedColumn_Success()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["MyColumn", "MyColumn"], heading.ColumnNames);
        Assert.Equal(0, heading.GetColumnIndex("MyColumn"));
        Assert.Equal(0, heading.GetColumnIndex("mycolumn"));
    }

    [Fact]
    public void GetColumnIndex_DuplicatedColumnEmpty_Success()
    {
        using var importer = Helpers.GetImporter("EmptyColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["", "Column2", "", " Column4 "], heading.ColumnNames);
        Assert.Equal(0, heading.GetColumnIndex(""));
        Assert.Equal(1, heading.GetColumnIndex("Column2"));
        Assert.Equal(3, heading.GetColumnIndex(" Column4 "));
    }

    [Fact]
    public void GetColumnIndex_NullColumnName_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ArgumentNullException>("columnName", () => heading.GetColumnIndex(null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("NoSuchColumn")]
    public void GetColumnIndex_NoSuchColumnName_ThrowsExcelMappingException(string columnName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => heading.GetColumnIndex(columnName));
    }

    [Fact]
    public void GetColumnName_GetFirstColumnMatchingIndex_Roundtrips()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        string[] columnNames = [.. heading.ColumnNames];
        for (int i = 0; i < columnNames.Length; i++)
        {
            var scopedIndex = i;
            Assert.Equal(i, heading.GetFirstColumnMatchingIndex(e => e == columnNames[scopedIndex]));
            Assert.Equal(columnNames[i], heading.GetColumnName(i));
        }
    }

    [Fact]
    public void GetColumnName_NullPredicate_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ArgumentNullException>("predicate", () => heading.GetFirstColumnMatchingIndex(null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("NoSuchColumn")]
    public void GetFirstColumnMatchingIndex_NoSuchColumnMatching_ThrowsExcelMappingException(string columnName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => heading.GetFirstColumnMatchingIndex(e => e == columnName));
    }

    [Fact]
    public void TryGetColumnIndex_Invoke_Roundtrips()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        string[] columnNames = [.. heading.ColumnNames];
        for (int i = 0; i < columnNames.Length; i++)
        {
            Assert.True(heading.TryGetColumnIndex(columnNames[i], out int index));
            Assert.Equal(i, index);
            Assert.Equal(columnNames[i], heading.GetColumnName(i));
        }
    }

    [Fact]
    public void TryGetColumnIndex_InvokeValidColumnName_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.True(heading.TryGetColumnIndex("StringValue", out int index));
        Assert.Equal(1, index);
        Assert.True(heading.TryGetColumnIndex("stringvalue", out index));
        Assert.Equal(1, index);
        Assert.Equal(1, index);
    }

    [Fact]
    public void TryGetColumnIndex_DuplicatedColumn_Success()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["MyColumn", "MyColumn"], heading.ColumnNames);
        Assert.True(heading.TryGetColumnIndex("MyColumn", out int index));
        Assert.Equal(0, index);
        Assert.True(heading.TryGetColumnIndex("mycolumn", out index));
        Assert.Equal(0, index);
    }

    [Fact]
    public void TryGetColumnIndex_DuplicatedColumnEmpty_Success()
    {
        using var importer = Helpers.GetImporter("EmptyColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["", "Column2", "", " Column4 "], heading.ColumnNames);
        Assert.True(heading.TryGetColumnIndex("", out int index));
        Assert.Equal(0, index);
        Assert.True(heading.TryGetColumnIndex("Column2", out index));
        Assert.Equal(1, index);
        Assert.True(heading.TryGetColumnIndex(" Column4 ", out index));
        Assert.Equal(3, index);
    }

    [Fact]
    public void TryGetColumnIndex_NullColumnName_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ArgumentNullException>("columnName", () => heading.TryGetColumnIndex(null!, out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData("NoSuchColumn")]
    public void TryGetColumnIndex_NoSuchColumnName_ReturnsFalse(string columnName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.False(heading.TryGetColumnIndex(columnName, out int index));
        Assert.Equal(0, index);
    }

    [Fact]
    public void GetColumnName_TryGetFirstColumnMatchingIndex_Roundtrips()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        string[] columnNames = [.. heading.ColumnNames];
        for (int i = 0; i < columnNames.Length; i++)
        {
            var scopedIndex = i;
            Assert.True(heading.TryGetFirstColumnMatchingIndex(e => e == columnNames[scopedIndex], out int index));
            Assert.Equal(i, index);
            Assert.Equal(columnNames[i], heading.GetColumnName(i));
        }
    }

    [Fact]
    public void TryGetFirstColumnMatchingIndex_NullPredicate_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ArgumentNullException>("predicate", () => heading.TryGetFirstColumnMatchingIndex(null!, out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData("NoSuchColumn")]
    public void TryGetFirstColumnMatchingIndex_NoSuchColumnName_ReturnsFalse(string columnName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.False(heading.TryGetFirstColumnMatchingIndex(e => e == columnName, out int index));
        Assert.Equal(-1, index);
    }
}
