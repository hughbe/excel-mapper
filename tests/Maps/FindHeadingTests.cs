using System.Linq;

namespace ExcelMapper.Tests;

public class FindHeadingTests
{
    [Fact]
    public void FindHeading_TitleAndBlankRowsAbove_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Staff report,\nGenerated 2024-03-04,\n,\nName,Age\nAlice,34\nBob,45\n");
        var sheet = importer.ReadSheet();
        var heading = sheet.FindHeading<Person>();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["Name", "Age"], heading.ColumnNames);
        Assert.Equal(3, sheet.HeadingIndex);
        Assert.Equal(3, sheet.CurrentRowIndex);

        var rows = sheet.ReadRows<Person>().ToArray();
        Assert.Equal(["Alice", "Bob"], rows.Select(r => r.Name));
        Assert.Equal([34, 45], rows.Select(r => r.Age));
    }

    [Fact]
    public void FindHeading_Xlsx_ReturnsExpected()
    {
        using var importer = XlsxBuilder.Create(["Quarterly figures"], [null], ["Notes: provisional"], ["Name", "Age"], ["Alice", 34.0]);
        var sheet = importer.ReadSheet();
        sheet.FindHeading<Person>();
        Assert.Equal(3, sheet.HeadingIndex);
        var row = sheet.ReadRows<Person>().Single();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(34, row.Age);
    }

    [Fact]
    public void FindHeading_HeadingInFirstRow_ReturnsFirstRow()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,34\n");
        var sheet = importer.ReadSheet();
        sheet.FindHeading<Person>();
        Assert.Equal(0, sheet.HeadingIndex);
        Assert.Equal("Alice", sheet.ReadRows<Person>().Single().Name);
    }

    [Fact]
    public void FindHeading_RowWithSomeColumns_SkipsIt()
    {
        // A row with only some of the required columns isn't the heading.
        using var importer = Helpers.GetCsvImporter("Name,Comment\nName,Age\nAlice,34\n");
        var sheet = importer.ReadSheet();
        sheet.FindHeading<Person>();
        Assert.Equal(1, sheet.HeadingIndex);
        Assert.Equal(34, sheet.ReadRows<Person>().Single().Age);
    }

    public class Person
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }

    [Fact]
    public void FindHeading_OptionalColumns_NotRequired()
    {
        using var importer = Helpers.GetCsvImporter("Title\nName,Other\nAlice,x\n");
        var sheet = importer.ReadSheet();
        sheet.FindHeading<OptionalAgePerson>();
        Assert.Equal(1, sheet.HeadingIndex);
        var row = sheet.ReadRows<OptionalAgePerson>().Single();
        Assert.Equal("Alice", row.Name);
        Assert.Null(row.Age);
    }

    [Fact]
    public void FindHeading_OnlyOptionalColumns_NeedsOneOfThem()
    {
        // Every row has the required columns (there are none), so the heading is the first row
        // with any of the columns.
        using var importer = Helpers.GetCsvImporter("Report\nAge\n34\n");
        var sheet = importer.ReadSheet();
        sheet.FindHeading<OnlyOptional>();
        Assert.Equal(1, sheet.HeadingIndex);
        Assert.Equal(34, sheet.ReadRows<OnlyOptional>().Single().Age);
    }

    public class OptionalAgePerson
    {
        public string? Name { get; set; }

        [ExcelOptional]
        public int? Age { get; set; }
    }

    public class OnlyOptional
    {
        [ExcelOptional]
        public int? Age { get; set; }
    }

    [Fact]
    public void FindHeading_ColumnAttributesAndNestedObjects_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Customers\n\nFull Name,Street,Town,Code\nAl,High St,Leeds,7\n");
        var sheet = importer.ReadSheet();
        sheet.FindHeading<Customer>();
        Assert.Equal(2, sheet.HeadingIndex);
        var row = sheet.ReadRows<Customer>().Single();
        Assert.Equal("Al", row.Name);
        Assert.Equal("Leeds", row.Address.Town);
        Assert.Equal("7", row.Code);
    }

    public class Customer
    {
        [ExcelColumnName("Full Name")]
        public string? Name { get; set; }

        public Address Address { get; set; } = new();

        // Found by position, so it doesn't identify the heading.
        [ExcelColumnIndex(3)]
        public string? Code { get; set; }
    }

    public class Address
    {
        public string? Street { get; set; }
        public string? Town { get; set; }
    }

    [Fact]
    public void FindHeading_ClassMap_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Export\nRef,Amount\nX,10\n");
        importer.Configuration.RegisterClassMap<Order>(c =>
        {
            c.Map(o => o.Id).WithColumnName("Ref");
            c.Map(o => o.Total).WithColumnName("Amount");
        });
        var sheet = importer.ReadSheet();
        sheet.FindHeading<Order>();
        Assert.Equal(1, sheet.HeadingIndex);
        Assert.Equal(10, sheet.ReadRows<Order>().Single().Total);
    }

    public class Order
    {
        public string? Id { get; set; }
        public decimal Total { get; set; }
    }

    [Fact]
    public void FindHeading_IgnoreSeparatorsInColumnNames_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Title\nFirst Name,Last_Name\nAda,Lovelace\n");
        importer.Configuration.IgnoreSeparatorsInColumnNames = true;
        var sheet = importer.ReadSheet();
        sheet.FindHeading<Names>();
        Assert.Equal("Lovelace", sheet.ReadRows<Names>().Single().LastName);
    }

    public class Names
    {
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
    }

    [Fact]
    public void FindHeading_DictionaryOfAllColumns_FoundByOtherColumns()
    {
        using var importer = Helpers.GetCsvImporter("Title\nName,Age\nAlice,34\n");
        var sheet = importer.ReadSheet();
        sheet.FindHeading<WithExtra>();
        var row = sheet.ReadRows<WithExtra>().Single();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(new Dictionary<string, string> { ["Age"] = "34" }, row.Extra);
    }

    public class WithExtra
    {
        public string? Name { get; set; }

        [ExcelExtraColumns]
        public Dictionary<string, string> Extra { get; set; } = [];
    }

    [Fact]
    public void FindHeading_SplitListAndIndexers_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Title\nTags,First,Second\na;b,1,2\n");
        importer.Configuration.RegisterClassMap<Collections>(c =>
        {
            c.Map(p => p.Tags).WithSeparators(';');
            c.Map(p => p.Values[0]).WithColumnName("First");
            c.Map(p => p.Values[1]).WithColumnName("Second");
        });
        var sheet = importer.ReadSheet();
        sheet.FindHeading<Collections>();
        Assert.Equal(1, sheet.HeadingIndex);
        var row = sheet.ReadRows<Collections>().Single();
        Assert.Equal(["a", "b"], row.Tags);
        Assert.Equal([1, 2], row.Values);
    }

    [Fact]
    public void FindHeading_IndexerMissingColumn_SkipsRow()
    {
        using var importer = Helpers.GetCsvImporter("First\nFirst,Second\n1,2\n");
        importer.Configuration.RegisterClassMap<Collections>(c =>
        {
            c.Map(p => p.Values[0]).WithColumnName("First");
            c.Map(p => p.Values[1]).WithColumnName("Second");
        });
        var sheet = importer.ReadSheet();
        sheet.FindHeading<Collections>();
        Assert.Equal(1, sheet.HeadingIndex);
    }

    public class Collections
    {
        public string[] Tags { get; set; } = [];
        public int[] Values { get; set; } = new int[2];
    }

    [Fact]
    public void FindHeading_StartsAtHeadingIndex()
    {
        // Row 2 would match, but the search starts at HeadingIndex.
        using var importer = Helpers.GetCsvImporter("Title\nName,Age\nx\nName,Age\nAlice,34\n");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 2;
        sheet.FindHeading<Person>();
        Assert.Equal(3, sheet.HeadingIndex);
        Assert.Equal("Alice", sheet.ReadRows<Person>().Single().Name);
    }

    [Fact]
    public void FindHeading_HasHeadingFalse_SetsHasHeading()
    {
        using var importer = Helpers.GetCsvImporter("Title\nName,Age\nAlice,34\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        sheet.FindHeading<Person>();
        Assert.True(sheet.HasHeading);
        Assert.Equal("Alice", sheet.ReadRows<Person>().Single().Name);
    }

    [Fact]
    public void FindHeading_AfterReadingRows_SearchesFromStart()
    {
        using var importer = Helpers.GetCsvImporter("Title\nName,Age\nAlice,34\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        Assert.Equal("Title", sheet.ReadRow<Line>().Text);

        sheet.FindHeading<Person>();
        Assert.Equal(1, sheet.HeadingIndex);
        Assert.Equal("Alice", sheet.ReadRows<Person>().Single().Name);
    }

    public class Line
    {
        [ExcelColumnIndex(0)]
        public string? Text { get; set; }
    }

    [Fact]
    public void FindHeading_NotFound_ThrowsWithClosestRow()
    {
        using var importer = Helpers.GetCsvImporter("Title\nName,Agee\nAlice,34\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.FindHeading<Person>());
        Assert.Contains("The closest is row 2", exception.Message);
        Assert.Contains("\"Age\"", exception.Message);
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void FindHeading_NoRowHasAnyColumn_Throws()
    {
        using var importer = Helpers.GetCsvImporter("a,b\nc,d\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.FindHeading<Person>());
        Assert.Contains("none of the first 50 rows", exception.Message);
    }

    [Fact]
    public void FindHeading_BeyondMaxRows_Throws()
    {
        using var importer = Helpers.GetCsvImporter("1\n2\n3\n4\nName,Age\nAlice,34\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.FindHeading<Person>(4));
        sheet.FindHeading<Person>(5);
        Assert.Equal(4, sheet.HeadingIndex);
        Assert.Equal("Alice", sheet.ReadRows<Person>().Single().Name);
    }

    [Theory]
    [InlineData(49, true)]
    [InlineData(50, false)]
    public void FindHeading_Default_SearchesFirst50Rows(int rowsAbove, bool found)
    {
        var csv = string.Concat(Enumerable.Repeat("x\n", rowsAbove)) + "Name,Age\nAlice,34\n";
        using var importer = Helpers.GetCsvImporter(csv);
        var sheet = importer.ReadSheet();
        Assert.Equal(found, sheet.TryFindHeading<Person>(out _));
        if (found)
        {
            Assert.Equal(rowsAbove, sheet.HeadingIndex);
            Assert.Equal("Alice", sheet.ReadRows<Person>().Single().Name);
        }
    }

    [Fact]
    public void TryFindHeading_NotFound_LeavesSheetUnchanged()
    {
        using var importer = Helpers.GetCsvImporter("Title\nName,Agee\nAlice,34\n");
        var sheet = importer.ReadSheet();
        Assert.False(sheet.TryFindHeading<Person>(out var heading));
        Assert.Null(heading);
        Assert.Null(sheet.Heading);
        Assert.Equal(0, sheet.HeadingIndex);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading starts from the beginning again.
        sheet.HeadingIndex = 1;
        Assert.Equal("Alice", sheet.ReadRows<OptionalAgePerson>().Single().Name);
    }

    [Fact]
    public void TryFindHeading_Found_ReturnsTrue()
    {
        using var importer = Helpers.GetCsvImporter("Title\nName,Age\nAlice,34\n");
        var sheet = importer.ReadSheet();
        Assert.True(sheet.TryFindHeading<Person>(10, out var heading));
        Assert.Same(sheet.Heading, heading);
        Assert.Equal(1, sheet.HeadingIndex);
    }

    [Fact]
    public void FindHeading_NoNamedColumns_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Title\nA,B\n1,2\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.FindHeading<Line>());
        Assert.Contains("doesn't look up any column by name", exception.Message);
        Assert.Throws<ExcelMappingException>(() => sheet.TryFindHeading<Line>(out _));
    }

    [Fact]
    public void FindHeading_ColumnsByOtherNames_ReturnsExpected()
    {
        // Columns found by any of several names, by a list of names, or by a pattern are looked
        // up by name, so they identify the heading row.
        using var importer = Helpers.GetCsvImporter("Report,,,\nFull Name,Q1,Q2,Total\nAlice,1,2,3\n");
        var sheet = importer.ReadSheet();
        sheet.FindHeading<OtherNames>();
        Assert.Equal(1, sheet.HeadingIndex);

        var row = sheet.ReadRows<OtherNames>().Single();
        Assert.Equal("Alice", row.Name);
        Assert.Equal([1, 2], row.Quarters);
        Assert.Equal(3, row.Total);
    }

    private class OtherNames
    {
        [ExcelColumnName("Name")]
        [ExcelColumnName("Full Name")]
        public string Name { get; set; } = default!;

        [ExcelColumnNames("Q1", "Q2")]
        public int[] Quarters { get; set; } = default!;

        [ExcelColumnMatching("^Tot")]
        public int Total { get; set; }
    }
}
