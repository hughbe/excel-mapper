using System.Linq;

namespace ExcelMapper.Tests;

public class MapDuplicatedColumnsTests
{
    [Fact]
    public void ReadHeading_WithDuplicatedColumns_KeepsBothNames()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Equal(["MyColumn", "MyColumn"], heading.ColumnNames.ToArray());
    }

    [Fact]
    public void ReadRow_AutoMapped_ThrowsExcelMappingException()
    {
        // It isn't clear which column to read, so neither is read.
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MyValuesWithOneColumn>());
        Assert.Equal("More than one column matches \"MyColumn\" (columns A and B), so it isn't clear which to read. Rename one of the columns, or choose one by its letter or index (ExcelColumnLetter or WithColumnLetter) in sheet \"Primitives\".", exception.Message);
        Assert.Equal("MyColumn", exception.ColumnName);
        Assert.Same(sheet, exception.Sheet);
    }

    [Fact]
    public void ReadRow_ColumnName_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        importer.Configuration.RegisterClassMap<MyValues>(c =>
        {
            c.Map(o => o.MyColumn).WithColumnName("MyColumn");
            c.Map(o => o.AnotherColumn).WithColumnIndex(1);
        });

        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRow<MyValues>());
    }

    [Fact]
    public void ReadRow_ColumnIndices_ReadsEachColumn()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        importer.Configuration.RegisterClassMap<MyValues>(c =>
        {
            c.Map(o => o.MyColumn).WithColumnIndex(0);
            c.Map(o => o.AnotherColumn).WithColumnLetter("B");
        });

        var row1 = importer.ReadSheet().ReadRow<MyValues>();
        Assert.Equal("value1", row1.MyColumn);
        Assert.Equal("value2", row1.AnotherColumn);
    }

    [Fact]
    public void ReadRows_DuplicatedColumnNotMapped_ReadsOtherColumns()
    {
        using var importer = Helpers.GetCsvImporter("Name,Notes,Notes\nAlice,a,b\n");
        Assert.Equal("Alice", importer.ReadSheet().ReadRows<NameRow>().Single().Name);
    }

    [Fact]
    public void ReadRows_DuplicatedBlankHeadings_ReadsNamedColumns()
    {
        using var importer = Helpers.GetCsvImporter("Name,,\nAlice,a,b\n");
        Assert.Equal("Alice", importer.ReadSheet().ReadRows<NameRow>().Single().Name);
    }

    [Fact]
    public void ReadRows_DuplicatedInAnotherCase_ThrowsExcelMappingException()
    {
        // Names are matched ignoring case, so "name" and "Name" both match.
        using var importer = Helpers.GetCsvImporter("name,Name\na,b\n");
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<NameRow>().ToArray());
        Assert.Contains("(columns A and B)", exception.Message);
    }

    [Fact]
    public void ReadRows_DuplicatedInAnotherCaseOrdinal_ReadsMatchingColumn()
    {
        // A case-sensitive name only matches one of them, even though the other comes first.
        using var importer = Helpers.GetCsvImporter("name,Name\na,b\n");
        Assert.Equal("b", importer.ReadSheet().ReadRows<OrdinalNameRow>().Single().Name);
    }

    [Fact]
    public void ReadRows_DuplicatedOrdinal_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("x,Name,Name\na,b,c\n");
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<OrdinalNameRow>().ToArray());
        Assert.Contains("(columns B and C)", exception.Message);
    }

    [Fact]
    public void ReadRows_ColumnNamesFirstDuplicated_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name,Name,Title\na,b,c\n");
        importer.Configuration.RegisterClassMap<NameRow>(c => c.Map(o => o.Name).WithColumnNames("Name", "Title"));
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<NameRow>().ToArray());
    }

    [Fact]
    public void ReadRows_IgnoreSeparatorsTwoColumnsMatch_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("First Name,First_Name\na,b\n");
        importer.Configuration.IgnoreSeparatorsInColumnNames = true;
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<FirstNameRow>().ToArray());
        Assert.Contains("\"FirstName\" (columns A and B)", exception.Message);
    }

    [Fact]
    public void ReadRows_IgnoreSeparatorsExactMatch_PrefersExactMatch()
    {
        // An exact match is preferred to one that ignores separators, so it isn't ambiguous.
        using var importer = Helpers.GetCsvImporter("First Name,FirstName\na,b\n");
        importer.Configuration.IgnoreSeparatorsInColumnNames = true;
        Assert.Equal("b", importer.ReadSheet().ReadRows<FirstNameRow>().Single().FirstName);
    }

    [Fact]
    public void ReadRowResults_Duplicated_ThrowsOnce()
    {
        // A problem with the heading, so it's reported once rather than as each row's error.
        using var importer = Helpers.GetCsvImporter("Name,Name\na,b\nc,d\n");
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRowResults<NameRow>());
        Assert.Contains("More than one column matches \"Name\" (columns A and B)", exception.Message);
    }

    [Fact]
    public void TryValidateHeading_Duplicated_ReturnsProblem()
    {
        using var importer = Helpers.GetCsvImporter("Name,Name\na,b\n");
        Assert.False(importer.ReadSheet().TryValidateHeading<NameRow>(out var problems));
        var problem = Assert.Single(problems);
        Assert.StartsWith("More than one column matches \"Name\" (columns A and B)", problem);
        Assert.EndsWith("(member Name)", problem);
    }

    [Fact]
    public void FindHeading_Duplicated_FindsHeading()
    {
        // The heading is found, and reading reports which column to choose.
        using var importer = Helpers.GetCsvImporter("Title\nName,Name\na,b\n");
        var sheet = importer.ReadSheet();
        sheet.FindHeading<NameRow>();
        Assert.Equal(1, sheet.HeadingIndex);
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NameRow>().ToArray());
    }

    private class MyValuesWithOneColumn
    {
        public string MyColumn { get; set; } = default!;
    }

    private class MyValues
    {
        public string MyColumn { get; set; } = default!;
        public string AnotherColumn { get; set; } = default!;
    }

    private class NameRow
    {
        public string Name { get; set; } = default!;
    }

    private class OrdinalNameRow
    {
        [ExcelColumnName("Name", StringComparison.Ordinal)]
        public string Name { get; set; } = default!;
    }

    private class FirstNameRow
    {
        public string FirstName { get; set; } = default!;
    }
}
