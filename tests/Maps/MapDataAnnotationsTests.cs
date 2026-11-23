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
}
