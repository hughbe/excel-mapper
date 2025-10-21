using System;
using System.Collections;
using Xunit;

namespace ExcelMapper.Tests;

public class AutoMapTests
{
    [Fact]
    public void ReadRow_AutoMappedEmptyClass_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyClass>();
        Assert.IsType<EmptyClass>(row1);

        // Valid value
        var row2 = sheet.ReadRow<EmptyClass>();
        Assert.IsType<EmptyClass>(row2);

        // Empty value
        var row3 = sheet.ReadRow<EmptyClass>();
        Assert.IsType<EmptyClass>(row3);

        // Last row.
        var row4 = sheet.ReadRow<EmptyClass>();
        Assert.IsType<EmptyClass>(row4);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyClass>());
    }
    
    private class EmptyClass
    {
    }
    
    [Fact]
    public void ReadRow_AutoMappedPropertyClass_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<SimplePropertyClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<SimplePropertyClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<SimplePropertyClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<SimplePropertyClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SimplePropertyClass>());
    }
    
    private class SimplePropertyClass
    {
        public string Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedFieldClass_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<SimpleFieldClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<SimpleFieldClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<SimpleFieldClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<SimpleFieldClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SimpleFieldClass>());
    }
    
    private class SimpleFieldClass
    {
        public string Value = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedReadonlyFieldClass_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ReadonlyFieldClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ReadonlyFieldClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ReadonlyFieldClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<ReadonlyFieldClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ReadonlyFieldClass>());
    }
    
    private class ReadonlyFieldClass
    {
        public readonly string Value = default!;
    }
    
    [Fact]
    public void ReadRow_AutoMappedClassWithRequiredProperty_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ClassWithRequiredProperty>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ClassWithRequiredProperty>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ClassWithRequiredProperty>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<ClassWithRequiredProperty>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ClassWithRequiredProperty>());
    }
    
    private class ClassWithRequiredProperty
    {
        public required string Value { get; set; } = default!;
    }
    
    [Fact]
    public void ReadRow_AutoMappedClassWithInitProperty_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ClassWithInitProperty>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ClassWithInitProperty>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ClassWithInitProperty>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<ClassWithInitProperty>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ClassWithInitProperty>());
    }
    
    private class ClassWithInitProperty
    {
        public string Value { get; init; } = default!;
    }
    
    [Fact]
    public void ReadRow_AutoMappedClassWithIndexerProperty_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ClassWithIndexerProperty>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ClassWithIndexerProperty>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ClassWithIndexerProperty>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<ClassWithIndexerProperty>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ClassWithIndexerProperty>());
    }
    
    private class ClassWithIndexerProperty
    {
        public string Value { get; set; } = default!;

        public string this[int index]
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }
    }

    [Fact]
    public void ReadRow_AutoMappedInterfacePropertyClass_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InterfacePropertyClass>());
    }

    private class InterfacePropertyClass
    {
        public IDisposable Value { get; set;  } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedInterfaceFieldClass_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InterfaceFieldClass>());
    }
    
    private class InterfaceFieldClass
    {
        public IDisposable Value = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedAbstractClass_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<AbstractClass>());
    }

    private abstract class AbstractClass
    {
    }

    [Fact]
    public void ReadRow_AutoMappedClassWithConstructor_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ClassWithConstructor>());
    }

    private class ClassWithConstructor
    {
        public string Value { get; set; } = default!;

        public ClassWithConstructor(string value)
        {
        }
    }

    [Fact]
    public void ReadRow_AutoMappedBitArray_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BitArray>());
    }
}
