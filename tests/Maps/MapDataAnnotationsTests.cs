using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelMapper.Tests;

public class MapDataAnnotationsTests
{
    [Fact]
    public void ReadRow_DataAnnotationsValidation_Success()
    {
        using var importer = Helpers.GetImporter("DataAnnotations.xlsx");
        importer.Configuration.ValidateDataAnnotations = true;

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ValidatedDataClass>();
        Assert.Equal(85, row1.Percentage);

        // Invalid cell value.
        Assert.IsType<ValidationException>(Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ValidatedDataClass>()).InnerException);
    }

    [Fact]
    public void ReadRows_InvalidDataAnnotation_ThrowsExcelMappingExceptionWithRowAndMember()
    {
        using var importer = Helpers.GetCsvImporter("Percentage\n50\n500\n");
        importer.Configuration.ValidateDataAnnotations = true;

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<ValidatedDataClass>().ToArray());
        Assert.Equal("The field Percentage must be between 1 and 100 on row 3.", exception.Message);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(nameof(ValidatedDataClass.Percentage), exception.Member!.Name);
        var inner = Assert.IsType<ValidationException>(exception.InnerException);
        Assert.Equal([nameof(ValidatedDataClass.Percentage)], inner.ValidationResult.MemberNames);
    }

    private class ValidatedDataClass
    {
        [Range(1, 100)]
        public int Percentage { get; set; }
    }

    [Fact]
    public void ReadRow_NoDataAnnotationsValidation_Success()
    {
        using var importer = Helpers.GetImporter("DataAnnotations.xlsx");
        importer.Configuration.ValidateDataAnnotations = false;

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ValidatedDataClass>();
        Assert.Equal(85, row1.Percentage);

        // Invalid cell value.
        var row2 = sheet.ReadRow<ValidatedDataClass>();
        Assert.Equal(150, row2.Percentage);
    }

    [Fact]
    public void ReadRow_DataAnnotationsValidationString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.ValidateDataAnnotations = true;

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<string>();
        Assert.Equal("value", row1);

        // Valid value
        var row2 = sheet.ReadRow<string>();
        Assert.Equal("  value  ", row2);

        // Empty value
        var row3 = sheet.ReadRow<string>();
        Assert.Null(row3);

        // Last row.
        var row4 = sheet.ReadRow<string>();
        Assert.Equal("value", row4);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<string>());
    }

    [Fact]
    public void ReadRow_DataAnnotationsValidationMultiple_Success()
    {
        using var importer = Helpers.GetImporter("DataAnnotations.xlsx");
        importer.Configuration.ValidateDataAnnotations = true;

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        Assert.IsType<ValidationException>(Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MultipleValidatedDataClass>()).InnerException);

        // Invalid cell value.
        Assert.IsType<ValidationException>(Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MultipleValidatedDataClass>()).InnerException);
    }

    private class MultipleValidatedDataClass
    {
        [Range(1, 100)]
        public int Percentage { get; set; }

        [ExcelIgnore]
        [Required]
        [StringLength(5, MinimumLength = 2)]
        public string? ShortText { get; set; }
    }

    [Fact]
    public void ReadRows_InvalidNestedObject_ThrowsExcelMappingExceptionNamingObject()
    {
        using var importer = Helpers.GetCsvImporter("Name,Street,Postcode\nAlice,,AB1\n");
        importer.Configuration.ValidateDataAnnotations = true;

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<Customer>().ToArray());
        Assert.Equal("The Street field is required in Address on row 2.", exception.Message);
        Assert.Equal(nameof(CustomerAddress.Street), exception.Member!.Name);
        Assert.IsType<ValidationException>(exception.InnerException);
    }

    [Fact]
    public void ReadRowResults_InvalidNestedObjects_ReportsEveryFailure()
    {
        using var importer = Helpers.GetCsvImporter("Name,Street,Postcode\n,,TOOLONGPOSTCODE\nBob,High Street,AB1\n");
        importer.Configuration.ValidateDataAnnotations = true;

        var results = importer.ReadSheet().ReadRowResults<Customer>().ToArray();
        Assert.Equal(
            [
                "The Name field is required on row 2.",
                "The Street field is required in Address on row 2.",
                "The field Postcode must be a string or array type with a maximum length of '8' in Address.Code on row 2.",
            ],
            results[0].Errors.Select(e => e.Message));
        Assert.True(results[1].IsValid);
    }

    [Fact]
    public void ReadRows_NestedObjectsValid_ReturnsRows()
    {
        using var importer = Helpers.GetCsvImporter("Name,Street,Postcode\nAlice,High Street,AB1\n");
        importer.Configuration.ValidateDataAnnotations = true;
        Assert.Equal("AB1", importer.ReadSheet().ReadRows<Customer>().Single().Address.Code.Postcode);
    }

    [Fact]
    public void ReadRows_NestedObjectNotMapped_IsntValidated()
    {
        // Only nested objects the class map reads are validated: an object the row doesn't set
        // isn't from the file.
        using var importer = Helpers.GetCsvImporter("Name\nAlice\n");
        importer.Configuration.ValidateDataAnnotations = true;
        importer.Configuration.RegisterClassMap<Customer>(c => c.Map(r => r.Name));

        var row = importer.ReadSheet().ReadRows<Customer>().Single();
        Assert.Null(row.Address.Street);
    }

    public class Customer
    {
        [Required]
        public string? Name { get; set; }

        public CustomerAddress Address { get; set; } = new();
    }

    public class CustomerAddress
    {
        [Required]
        public string? Street { get; set; }

        public PostcodeHolder Code { get; set; } = new();
    }

    public class PostcodeHolder
    {
        [MaxLength(8)]
        public string? Postcode { get; set; }
    }
}
