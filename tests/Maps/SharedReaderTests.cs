using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Every sheet of an importer shares one forward-only reader. Each sheet keeps its own position,
/// so reading from one sheet doesn't change what another reads next.
/// </summary>
public class SharedReaderTests
{
    [Fact]
    public void ReadRow_TwoSheetObjectsForSameWorksheet_EachKeepsItsPosition()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\nc\nd\n");
        var first = importer.ReadSheet("");
        var second = importer.ReadSheet("");

        Assert.Equal("a", first.ReadRow<Row>().Name);
        Assert.Equal("b", first.ReadRow<Row>().Name);
        Assert.Equal("a", second.ReadRow<Row>().Name);
        Assert.Equal("c", first.ReadRow<Row>().Name);
        Assert.Equal("b", second.ReadRow<Row>().Name);
        Assert.Equal("d", first.ReadRow<Row>().Name);
        Assert.False(first.TryReadRow<Row>(out _));
        Assert.Equal("c", second.ReadRow<Row>().Name);
        Assert.Equal(4, first.CurrentRowIndex);
        Assert.Equal(3, second.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_InterleavedEnumeratorsForSameWorksheet_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\nc\n");
        var first = importer.ReadSheet("").ReadRows<Row>();
        var second = importer.ReadSheet("").ReadRows<Row>();
        var pairs = first.Zip(second, (x, y) => $"{x.Name}{y.Name}").ToArray();
        Assert.Equal(["aa", "bb", "cc"], pairs);
    }

    [Fact]
    public void ReadRow_DifferentWorksheetsInterleaved_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var primitives = importer.ReadSheet("Primitives");
        var third = importer.ReadSheet("Third Sheet");

        Assert.Equal("a", primitives.ReadRow<Row>().StringValue);
        Assert.Equal("s1", third.ReadRow<Row>().StringValue);
        Assert.Equal("b", primitives.ReadRow<Row>().StringValue);
        Assert.Equal("s2", third.ReadRow<Row>().StringValue);
        Assert.Equal(2, primitives.CurrentRowIndex);
        Assert.Equal(2, third.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_RangeOnOneSheetObject_DoesNotMoveAnother()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\nc\nd\n");
        var sheet = importer.ReadSheet("");
        var other = importer.ReadSheet("");
        Assert.Equal("a", sheet.ReadRow<Row>().Name);

        Assert.Equal(["c", "d"], other.ReadRows<Row>(3, 2).Select(r => r.Name));
        Assert.Equal("b", sheet.ReadRow<Row>().Name);
    }

    [Fact]
    public void FindHeading_WhileAnotherSheetObjectIsReading_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Title,\nName,Age\na,1\nb,2\n");
        var first = importer.ReadSheet("");
        first.HeadingIndex = 1;
        Assert.Equal("a", first.ReadRow<Row>().Name);

        var second = importer.ReadSheet("");
        second.FindHeading<Person>();
        Assert.Equal(["a", "b"], second.ReadRows<Person>().Select(p => p.Name));
        Assert.Equal("b", first.ReadRow<Row>().Name);
    }

    [Fact]
    public void ReadRows_CsvDataRangeFromEndWhileAnotherSheetObjectIsReading_ReturnsExpected()
    {
        // Counting the rows of a CSV file reads to its end; other sheet objects continue from
        // where they were.
        using var importer = Helpers.GetCsvImporter("Name\na\n\nb\nTotal\n");
        var first = importer.ReadSheet("");
        Assert.Equal("a", first.ReadRow<Row>().Name);

        var second = importer.ReadSheet("");
        second.DataRange = new ExcelRange(..^1, Range.All);
        Assert.Equal(["a", null, "b"], second.ReadRows<Row>().Select(r => r.Name));

        Assert.Null(first.ReadRow<Row>().Name);
        Assert.Equal("b", first.ReadRow<Row>().Name);
        Assert.Equal("Total", first.ReadRow<Row>().Name);
    }

    [Fact]
    public void ReadSheet_AfterReadingRows_ContinuesFromThatSheet()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var third = importer.ReadSheet("Third Sheet");
        Assert.Equal("s1", third.ReadRow<Row>().StringValue);

        // Sheets are read in order from the sheet whose rows were read last.
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet());

        var primitives = importer.ReadSheet(0);
        Assert.Equal("a", primitives.ReadRow<Row>().StringValue);
        Assert.Equal("Empty", importer.ReadSheet().Name);
        Assert.Equal("s2", third.ReadRow<Row>().StringValue);
    }

    [Fact]
    public void ReadRow_TrailingEmptyLines_CurrentRowIndexIsLastRow()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\n\n\n");
        var sheet = importer.ReadSheet();
        Assert.Equal(["a", "b"], sheet.ReadRows<Row>().Select(r => r.Name));
        Assert.Equal(2, sheet.CurrentRowIndex);
        Assert.False(sheet.TryReadRow<Row>(out _));
        Assert.Equal(2, sheet.CurrentRowIndex);
    }

    public class Row
    {
        [ExcelOptional]
        public string? Name { get; set; }

        [ExcelOptional]
        public string? StringValue { get; set; }
    }

    public class Person
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }
}
