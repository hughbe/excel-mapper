using System.Linq;
namespace ExcelMapper.Tests;
 
public class ExcelMappingExceptionTests
{
    [Fact]
    public void Ctor_Default()
    {
        var exception = new ExcelMappingException();
        Assert.NotNull(exception.Message);
        Assert.Null(exception.InnerException);
        Assert.Null(exception.Sheet);
        Assert.Equal(-1, exception.RowIndex);
        Assert.Equal(-1, exception.ColumnIndex);
    }

    [Fact]
    public void Ctor_Message()
    {
        var exception = new ExcelMappingException("message");
        Assert.Equal("message", exception.Message);
        Assert.Null(exception.InnerException);
        Assert.Null(exception.Sheet);
        Assert.Equal(-1, exception.RowIndex);
        Assert.Equal(-1, exception.ColumnIndex);
    }

    [Fact]
    public void Ctor_Message_InnerException()
    {
        var innerException = new DivideByZeroException();
        var exception = new ExcelMappingException("message", innerException);
        Assert.Equal("message", exception.Message);
        Assert.Same(innerException, exception.InnerException);
        Assert.Null(exception.Sheet);
        Assert.Equal(-1, exception.RowIndex);
        Assert.Equal(-1, exception.ColumnIndex);
    }

    [Fact]
    public void Ctor_Message_SheetWithReadHeading_RowIndex_ColumnIndex()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var exception = new ExcelMappingException("Message", sheet, 10, 1);
        Assert.Equal("Message in column \"StringValue\" at cell B11 in sheet \"Primitives\".", exception.Message);
        Assert.Null(exception.InnerException);
        Assert.Same(sheet, exception.Sheet);
        Assert.Equal(10, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    [Fact]
    public void Ctor_Message_SheetWithNonReadHeading_RowIndex_ColumnIndex()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var exception = new ExcelMappingException("Message", sheet, 10, 1);

        // Creating an exception doesn't read the heading, as that would move the reader.
        Assert.Equal("Message at cell B11 in sheet \"Primitives\".", exception.Message);
        Assert.Null(sheet.Heading);
        Assert.Null(exception.ColumnName);
        Assert.Null(exception.InnerException);
        Assert.Same(sheet, exception.Sheet);
        Assert.Equal(10, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    [Fact]
    public void Ctor_Message_SheetNoHeading_RowIndex_ColumnIndex()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var exception = new ExcelMappingException("Message", sheet, 0, 27);
        Assert.Equal("Message at cell AB1 in sheet \"Primitives\".", exception.Message);
        Assert.Equal("AB1", exception.CellAddress);
    }

    [Fact]
    public void Ctor_Message_SheetColumnOutsideHeading_ReportsCellAddress()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var exception = new ExcelMappingException("Message", sheet, 4, 100);
        Assert.Equal("Message at cell CW5 in sheet \"Primitives\".", exception.Message);
        Assert.Equal("CW5", exception.CellAddress);
    }

    [Fact]
    public void Ctor_Message_SheetNoColumn_ReportsOneBasedRow()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();

        var exception = new ExcelMappingException("Message", sheet, 0, -1);
        Assert.Equal("Message on row 1 in sheet \"Primitives\".", exception.Message);
        Assert.Null(exception.CellAddress);
    }

    [Fact]
    public void CellAddress_FromReadRows_IsExcelAddress()
    {
        using var importer = Helpers.GetCsvImporter("Junk,Value\n1,2\n3,invalid\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntValueRow>().ToArray());
        Assert.Equal("B3", exception.CellAddress);
        Assert.Contains("in column \"Value\" at cell B3 in sheet", exception.Message);
    }

    private class IntValueRow
    {
        public int Value { get; set; }
    }

    [Fact]
    public void CellAddress_Default_IsNull()
    {
        Assert.Null(new ExcelMappingException().CellAddress);
        Assert.Null(new ExcelMappingException("message").CellAddress);
    }

    [Fact]
    public void Ctor_Default_MemberAndColumnNameAreNull()
    {
        var exception = new ExcelMappingException("message");
        Assert.Null(exception.Member);
        Assert.Null(exception.ColumnName);
    }

    [Fact]
    public void Ctor_SheetWithReadHeading_SetsColumnName()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Equal("StringValue", new ExcelMappingException("Message", sheet, 10, 1).ColumnName);
        Assert.Null(new ExcelMappingException("Message", sheet, 10, -1).ColumnName);
        Assert.Null(new ExcelMappingException("Message", sheet, 10, 100).ColumnName);
    }

    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    [Fact]
    public void ReadRows_InvalidValue_SetsMemberAndColumnName()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,old\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Person>().ToArray());
        Assert.Equal(nameof(Person.Age), exception.Member!.Name);
        Assert.Equal(typeof(Person), exception.Member.DeclaringType);
        Assert.Equal("Age", exception.ColumnName);
        Assert.Equal("B2", exception.CellAddress);
    }

    [Fact]
    public void ReadRows_MissingColumn_SetsMemberAndColumnName()
    {
        using var importer = Helpers.GetCsvImporter("Name\nAlice\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Person>().ToArray());
        Assert.Equal(nameof(Person.Age), exception.Member!.Name);
        Assert.Equal("Age", exception.ColumnName);
        Assert.Null(exception.CellAddress);
    }

    [Fact]
    public void ReadRows_EmptyValueType_SetsMemberAndColumnName()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Person>().ToArray());
        Assert.Equal(nameof(Person.Age), exception.Member!.Name);
        Assert.Equal("Age", exception.ColumnName);
    }

    [Fact]
    public void ReadRowResults_RecordedErrors_SetMemberAndColumnName()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,old\n");
        var error = Assert.Single(importer.ReadSheet().ReadRowResults<Person>().Single().Errors);
        Assert.Equal(nameof(Person.Age), error.Member!.Name);
        Assert.Equal("Age", error.ColumnName);
        Assert.Equal("B2", error.CellAddress);
    }
}
