using System.Linq;
using System.Collections;

namespace ExcelMapper.Tests;

public class AutoMapTests
{
    [Fact]
    public void ReadRow_AutoMappedEmptyClass_ThrowsExcelMappingException()
    {
        // Nothing in the type can be mapped, so every row would be an empty object.
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyClass>());
        Assert.StartsWith("The class map for type", exception.Message);
    }
    
    [Fact]
    public void ReadRows_AutoMappedEmptyClass_MessageNamesType()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<EmptyClass>().ToArray());
        Assert.Equal($"The class map for type \"{typeof(EmptyClass)}\" maps no members, so every row would be an empty object. Auto-mapping maps public fields, and public properties with a public setter or set by a public constructor; map other members with Map(...) in a class map.", exception.Message);
    }

    [Fact]
    public void ReadRows_AutoMappedPrivateSetters_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<PrivateSetterClass>().ToArray());
    }

    [Fact]
    public void ReadRows_AutoMappedEveryMemberIgnored_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<IgnoredClass>().ToArray());
    }

    [Fact]
    public void ReadRows_AutoMappedNestedEmptyClass_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<NestedEmptyClass>().ToArray());
        Assert.Contains($"\"{typeof(EmptyClass)}\"", exception.Message);
    }

    [Fact]
    public void ReadRows_RegisteredEmptyClassMap_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        importer.Configuration.RegisterClassMap<PublicSetterClass>(c => { });
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<PublicSetterClass>().ToArray());
    }

    [Fact]
    public void ReadRows_ClassMapWithPrivateSetterMapped_ReturnsExpected()
    {
        // A member auto-mapping skips can be mapped in a class map.
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        importer.Configuration.RegisterClassMap<PrivateSetterClass>(c => c.Map(p => p.Value));
        Assert.Equal(1, importer.ReadSheet().ReadRows<PrivateSetterClass>().Single().Value);
    }

    private class PrivateSetterClass
    {
        public int Value { get; private set; }
    }

    private class IgnoredClass
    {
        [ExcelIgnore]
        public int Value { get; set; }
    }

    private class NestedEmptyClass
    {
        public int Value { get; set; }
        public EmptyClass Nested { get; set; } = new();
    }

    private class PublicSetterClass
    {
        public int Value { get; set; }
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
    public void ReadRow_AutoMappedClassWithConstructor_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ClassWithConstructor>();
        Assert.Equal("value", row.Value);
        Assert.True(row.ConstructorCalled);
    }

    private class ClassWithConstructor
    {
        public string Value { get; set; } = default!;

        public bool ConstructorCalled { get; }

        public ClassWithConstructor(string value)
        {
            Value = value;
            ConstructorCalled = true;
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
