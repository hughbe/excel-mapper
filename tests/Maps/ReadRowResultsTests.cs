using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelMapper.Tests;

public class ReadRowResultsTests
{
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public DateTime Joined { get; set; }
    }

    private static readonly string s_joined = new DateTime(2024, 1, 2).ToString("G");

    [Fact]
    public void ReadRowResults_AllValid_ReturnsValues()
    {
        using var importer = Helpers.GetCsvImporter($"Name,Age,Joined\nAlice,30,{s_joined}\nBob,25,{s_joined}\n");
        var results = importer.ReadSheet().ReadRowResults<Person>().ToArray();

        Assert.Equal(2, results.Length);
        Assert.All(results, r => Assert.True(r.IsValid));
        Assert.All(results, r => Assert.Empty(r.Errors));
        Assert.Equal(["Alice", "Bob"], results.Select(r => r.Value!.Name));
        Assert.Equal([1, 2], results.Select(r => r.RowIndex));
    }

    [Fact]
    public void ReadRowResults_InvalidCells_ReportsEveryErrorAndContinues()
    {
        using var importer = Helpers.GetCsvImporter($"Name,Age,Joined\nAlice,30,{s_joined}\nBob,old,yesterday\nCarol,41,{s_joined}\n");
        var results = importer.ReadSheet().ReadRowResults<Person>().ToArray();

        Assert.Equal(3, results.Length);
        Assert.True(results[0].IsValid);
        Assert.True(results[2].IsValid);
        Assert.Equal("Carol", results[2].Value!.Name);

        var invalid = results[1];
        Assert.False(invalid.IsValid);
        Assert.Null(invalid.Value);
        Assert.Equal(2, invalid.RowIndex);
        Assert.Equal(["B3", "C3"], invalid.Errors.Select(e => e.CellAddress));
        Assert.Contains("\"old\"", invalid.Errors[0].Message);
        Assert.Contains("\"yesterday\"", invalid.Errors[1].Message);
    }

    [Fact]
    public void ReadRowResults_ThenReadRows_ThrowsAgain()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Joined\nBob,old,x\nBob,old,x\n");
        var sheet = importer.ReadSheet();

        Assert.False(sheet.ReadRowResults<Person>().First().IsValid);

        // Errors are only collected while enumerating ReadRowResults.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Person>());
    }

    public record PersonRecord(string Name, int Age)
    {
        public static int ConstructorCalls { get; set; }

        public string Name { get; } = Count(Name);

        private static string Count(string name)
        {
            ConstructorCalls++;
            return name;
        }
    }

    [Fact]
    public void ReadRowResults_RecordWithInvalidCell_DoesNotCallConstructor()
    {
        PersonRecord.ConstructorCalls = 0;
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,30\nBob,old\n");
        var results = importer.ReadSheet().ReadRowResults<PersonRecord>().ToArray();

        Assert.Equal(new PersonRecord("Alice", 30) with { }, results[0].Value);
        Assert.False(results[1].IsValid);
        Assert.Null(results[1].Value);
        Assert.Equal("B3", Assert.Single(results[1].Errors).CellAddress);
    }

    public class Validating
    {
        public Validating(int age)
        {
            if (age < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(age), "Age cannot be negative.");
            }

            Age = age;
        }

        public int Age { get; }
    }

    [Fact]
    public void ReadRowResults_ConstructorThrows_ReportsError()
    {
        using var importer = Helpers.GetCsvImporter("Age\n1\n-1\n2\n");
        var results = importer.ReadSheet().ReadRowResults<Validating>().ToArray();

        Assert.Equal([true, false, true], results.Select(r => r.IsValid));
        var error = Assert.Single(results[1].Errors);
        Assert.IsType<ArgumentOutOfRangeException>(error.InnerException);
        Assert.Contains("Age cannot be negative.", error.Message);
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
    }

    [Fact]
    public void ReadRowResults_NestedObjectError_IsReported()
    {
        using var importer = Helpers.GetCsvImporter("Name,City,Zip\nAlice,Paris,75001\nBob,Lyon,nope\n");
        importer.Configuration.RegisterClassMap<Customer>(c =>
        {
            c.Map(p => p.Name);
            c.Map(p => p.Address.City);
            c.Map(p => p.Address.Zip);
        });

        var results = importer.ReadSheet().ReadRowResults<Customer>().ToArray();
        Assert.True(results[0].IsValid);
        Assert.Equal(75001, results[0].Value!.Address.Zip);
        Assert.Equal("C3", Assert.Single(results[1].Errors).CellAddress);
    }

    [Fact]
    public void ReadRowResults_Primitive_ReportsInvalidRows()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\nx\n3\n");
        var results = importer.ReadSheet().ReadRowResults<int>().ToArray();

        Assert.Equal([true, false, true], results.Select(r => r.IsValid));
        Assert.Equal([1, 0, 3], results.Select(r => r.Value));
    }

    public class Annotated
    {
        [Range(0, 120)]
        public int Age { get; set; }

        [StringLength(3)]
        public string Code { get; set; } = "";
    }

    [Fact]
    public void ReadRowResults_DataAnnotations_ReportsEveryFailure()
    {
        using var importer = Helpers.GetCsvImporter("Age,Code\n30,ABC\n500,TOOLONG\n");
        importer.Configuration.ValidateDataAnnotations = true;
        var results = importer.ReadSheet().ReadRowResults<Annotated>().ToArray();

        Assert.True(results[0].IsValid);
        Assert.Equal(2, results[1].Errors.Count);
        Assert.All(results[1].Errors, e => Assert.IsType<ValidationException>(e.InnerException));
    }

    public class MissingColumn
    {
        public string Missing { get; set; } = "";
    }

    [Fact]
    public void ReadRowResults_MissingColumn_ThrowsOnce()
    {
        // A missing column is a problem with the sheet, not each row, so it's reported once,
        // when ReadRowResults is called, as ValidateHeading reports it.
        using var importer = Helpers.GetCsvImporter("Name\na\nb\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRowResults<MissingColumn>());
        Assert.Equal("The sheet is missing columns needed to map type ReadRowResultsTests.MissingColumn: column \"Missing\" (member Missing).", exception.Message);
    }

    [Fact]
    public void ReadRowResults_MissingColumnNoRows_Throws()
    {
        using var importer = Helpers.GetCsvImporter("Name\n");
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRowResults<MissingColumn>());
    }

    public class OptionalMissingColumn
    {
        public string Name { get; set; } = "";

        [ExcelOptional]
        public string Missing { get; set; } = "unchanged";
    }

    [Fact]
    public void ReadRowResults_OptionalMissingColumn_ReturnsRows()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\n");
        var results = importer.ReadSheet().ReadRowResults<OptionalMissingColumn>().ToArray();
        Assert.Equal(["a", "b"], results.Select(r => r.Value!.Name));
        Assert.All(results, r => Assert.Equal("unchanged", r.Value!.Missing));
    }

    public class Unmappable
    {
        public Func<int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRowResults_TypeCannotBeMapped_Throws()
    {
        using var importer = Helpers.GetCsvImporter("Value\na\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRowResults<Unmappable>().ToArray());
    }

    [Fact]
    public void ReadRowResults_ReaderClosed_Throws()
    {
        var importer = Helpers.GetCsvImporter("Value\n1\n");
        var sheet = importer.ReadSheet();
        importer.Dispose();
        Assert.Throws<ObjectDisposedException>(() => sheet.ReadRowResults<int>());
    }

    [Fact]
    public void ReadRowResults_ReadsHeadingImmediately()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var sheet = importer.ReadSheet();
        _ = sheet.ReadRowResults<int>();
        Assert.NotNull(sheet.Heading);
    }
}
