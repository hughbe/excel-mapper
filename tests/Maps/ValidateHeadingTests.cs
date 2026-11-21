using System.Linq;

namespace ExcelMapper.Tests;

public class ValidateHeadingTests
{
    public class Employee
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public decimal Salary { get; set; }

        [ExcelOptional]
        public string? Nickname { get; set; }
    }

    [Fact]
    public void ValidateHeading_AllColumnsPresent_Succeeds()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Salary\nAlice,30,100\n");
        var sheet = importer.ReadSheet();

        sheet.ValidateHeading<Employee>();
        Assert.True(sheet.TryValidateHeading<Employee>(out var problems));
        Assert.Empty(problems);

        // Validating doesn't read any rows.
        Assert.Equal("Alice", sheet.ReadRows<Employee>().Single().Name);
    }

    [Fact]
    public void ValidateHeading_MissingColumns_ReportsEveryMissingColumn()
    {
        using var importer = Helpers.GetCsvImporter("Name,Other\nAlice,x\n");
        var sheet = importer.ReadSheet();

        Assert.False(sheet.TryValidateHeading<Employee>(out var problems));
        Assert.Equal(["column \"Age\" (member Age)", "column \"Salary\" (member Salary)"], problems);

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ValidateHeading<Employee>());
        Assert.Equal($"Sheet \"\" is missing columns needed to map type \"{typeof(Employee)}\": column \"Age\" (member Age); column \"Salary\" (member Salary).", ex.Message);
    }

    [Fact]
    public void ValidateHeading_OptionalMember_NotRequired()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Salary\nAlice,30,100\n");
        Assert.True(importer.ReadSheet().TryValidateHeading<Employee>(out _));
    }

    [Fact]
    public void ValidateHeading_SimilarColumnName_SuggestsIt()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Annual_Salary,salary amount\nAlice,30,100,1\n");
        importer.Configuration.RegisterClassMap<Employee>(c =>
        {
            c.Map(p => p.Name);
            c.Map(p => p.Age);
            c.Map(p => p.Salary).WithColumnName("Salary Amount!");
        });

        Assert.False(importer.ReadSheet().TryValidateHeading<Employee>(out var problems));
        Assert.Equal("column \"Salary Amount!\" (member Salary)", Assert.Single(problems));

        using var importer2 = Helpers.GetCsvImporter("Name,Age,Sal-ary\nAlice,30,100\n");
        Assert.False(importer2.ReadSheet().TryValidateHeading<Employee>(out var problems2));
        Assert.Equal("column \"Salary\", did you mean \"Sal-ary\"? (member Salary)", Assert.Single(problems2));
    }

    public class Address
    {
        public string City { get; set; } = "";
        public int Zip { get; set; }
    }

    public class Customer
    {
        public string Name { get; set; } = "";
        public Address Address { get; set; } = default!;
        public string[] Tags { get; set; } = default!;
        public int[] Scores { get; set; } = default!;
    }

    [Fact]
    public void ValidateHeading_NestedCollectionsAndIndexers_ReportsPaths()
    {
        using var importer = Helpers.GetCsvImporter("Name,City,Tag1\nAlice,Paris,a\n");
        importer.Configuration.RegisterClassMap<Customer>(c =>
        {
            c.Map(p => p.Name);
            c.Map(p => p.Address.City);
            c.Map(p => p.Address.Zip);
            c.Map(p => p.Tags).WithColumnNames("Tag1", "Tag2");
            c.Map(p => p.Scores[0]).WithColumnName("Score1");
        });

        Assert.False(importer.ReadSheet().TryValidateHeading<Customer>(out var problems));
        Assert.Equal(
        [
            "column \"Zip\" (member Address.Zip)",
            "one of columns \"Tag1\", \"Tag2\" (member Tags)",
            "column \"Score1\" (member Scores[0])",
        ], problems);
    }

    public class Indexed
    {
        public string A { get; set; } = "";
        public string C { get; set; } = "";
    }

    [Fact]
    public void ValidateHeading_ColumnIndexBeyondSheet_ReportsColumnLetter()
    {
        using var importer = Helpers.GetCsvImporter("a,b\n1,2\n");
        importer.Configuration.RegisterClassMap<Indexed>(c =>
        {
            c.Map(p => p.A).WithColumnIndex(0);
            c.Map(p => p.C).WithColumnIndex(2);
        });

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        Assert.False(sheet.TryValidateHeading<Indexed>(out var problems));
        Assert.Equal("column C (index 2) (member C)", Assert.Single(problems));
    }

    [Fact]
    public void ValidateHeading_NamedColumnsWithoutHeading_ReportsProblem()
    {
        using var importer = Helpers.GetCsvImporter("Alice,30,100\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.False(sheet.TryValidateHeading<Employee>(out var problems));
        Assert.Equal(3, problems.Count);
        Assert.All(problems, p => Assert.Contains("does not have a heading", p));
    }

    public record Person(string Name, int Age);

    [Fact]
    public void ValidateHeading_Record_ReportsMissingParameterColumns()
    {
        using var importer = Helpers.GetCsvImporter("Name\nAlice\n");
        Assert.False(importer.ReadSheet().TryValidateHeading<Person>(out var problems));
        Assert.Equal("column \"Age\" (member Age)", Assert.Single(problems));
    }

    [Fact]
    public void ValidateHeading_Primitive_Succeeds()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        Assert.True(importer.ReadSheet().TryValidateHeading<int>(out _));
    }

    [Fact]
    public void ValidateHeading_ReaderClosed_Throws()
    {
        var importer = Helpers.GetCsvImporter("Value\n1\n");
        var sheet = importer.ReadSheet();
        importer.Dispose();
        Assert.Throws<ExcelMappingException>(() => sheet.ValidateHeading<int>());
    }
}
