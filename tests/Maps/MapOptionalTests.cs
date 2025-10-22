namespace ExcelMapper.Tests;

public class MapOptionalTests
{
    [Fact]
    public void ReadRows_AutoMappedPropertyDoesNotExist_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnPropertyClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedFieldDoesNotExist_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedAttributePropertyDoesNotExist_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnPropertyAttributeClass>();
        Assert.Equal(10, row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_AutoMappedAttributeFieldDoesNotExist_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnFieldAttributeClass>();
        Assert.Equal(10, row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_AutoMappedPropertyDoesNotExistEnumerable_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnPropertyEnumerableClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedFieldDoesNotExistEnumerable_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_AutoMappedAttributePropertyDoesNotExistEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnPropertyEnumerableAttributeClass>();
        Assert.Equal([10], row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_AutoMappedAttributeFieldDoesNotExistEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnEnumerableFieldAttributeClass>();
        Assert.Equal([10], row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_DefaultMappedPropertyDoesNotExist_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnPropertyClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedFieldDoesNotExist_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedAttributePropertyDoesNotExist_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnPropertyAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnPropertyAttributeClass>();
        Assert.Equal(10, row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_DefaultMappedAttributeFieldDoesNotExist_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnFieldAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnFieldAttributeClass>();
        Assert.Equal(10, row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_DefaultMappedPropertyDoesNotExistEnumerable_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnPropertyEnumerableClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnPropertyEnumerableClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedFieldDoesNotExistEnumerable_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnEnumerableFieldClass>());
    }

    [Fact]
    public void ReadRows_DefaultMappedAttributePropertyDoesNotExistEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnPropertyEnumerableAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnPropertyEnumerableAttributeClass>();
        Assert.Equal([10], row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_DefaultMappedAttributeFieldDoesNotExistEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnEnumerableFieldAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnEnumerableFieldAttributeClass>();
        Assert.Equal([10], row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_CustomMappedPropertyDoesNotExist_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomMissingColumnPropertyClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnPropertyClass>();
        Assert.Equal(10, row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_CustomMappedFieldDoesNotExist_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomMissingColumnFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnFieldClass>();
        Assert.Equal(10, row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_CustomMappedPropertyDoesNotExistEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomMissingColumnPropertyEnumerableClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnPropertyEnumerableAttributeClass>();
        Assert.Equal([10], row1.NoSuchColumn);
    }

    [Fact]
    public void ReadRows_CustomMappedFieldDoesNotExistEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<CustomMissingColumnEnumerableFieldClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnEnumerableFieldClass>();
        Assert.Equal([10], row1.NoSuchColumn);
    }

    private class MissingColumnPropertyClass
    {
        public int NoSuchColumn { get; set; } = 10;
    }

    private class DefaultMissingColumnPropertyClassMap : ExcelClassMap<MissingColumnPropertyClass>
    {
        public DefaultMissingColumnPropertyClassMap()
        {
            Map(p => p.NoSuchColumn);
        }
    }

    private class CustomMissingColumnPropertyClassMap : ExcelClassMap<MissingColumnPropertyClass>
    {
        public CustomMissingColumnPropertyClassMap()
        {
            Map(p => p.NoSuchColumn)
                .MakeOptional();
        }
    }

    private class MissingColumnPropertyAttributeClass
    {
        [ExcelOptional]
        public int NoSuchColumn { get; set; } = 10;
    }

    private class DefaultMissingColumnPropertyAttributeClassMap : ExcelClassMap<MissingColumnPropertyAttributeClass>
    {
        public DefaultMissingColumnPropertyAttributeClassMap()
        {
            Map(p => p.NoSuchColumn);
        }
    }

    private class CustomMissingColumnPropertyAttributeClassMap : ExcelClassMap<MissingColumnPropertyAttributeClass>
    {
        public CustomMissingColumnPropertyAttributeClassMap()
        {
            Map(p => p.NoSuchColumn)
                .MakeOptional();
        }
    }

    private class MissingColumnPropertyEnumerableClass
    {
        public int[] NoSuchColumn { get; set; } = [10];
    }

    private class DefaultMissingColumnPropertyEnumerableClassMap : ExcelClassMap<MissingColumnPropertyEnumerableClass>
    {
        public DefaultMissingColumnPropertyEnumerableClassMap()
        {
            Map(p => p.NoSuchColumn);
        }
    }

    private class CustomMissingColumnPropertyEnumerableClassMap : ExcelClassMap<MissingColumnPropertyEnumerableClass>
    {
        public CustomMissingColumnPropertyEnumerableClassMap()
        {
            Map(p => p.NoSuchColumn)
                .MakeOptional();
        }
    }

    private class MissingColumnPropertyEnumerableAttributeClass
    {
        [ExcelOptional]
        public int[] NoSuchColumn { get; set; } = [10];
    }

    private class DefaultMissingColumnPropertyEnumerableAttributeClassMap : ExcelClassMap<MissingColumnPropertyEnumerableAttributeClass>
    {
        public DefaultMissingColumnPropertyEnumerableAttributeClassMap()
        {
            Map(p => p.NoSuchColumn);
        }
    }

    private class CustomMissingColumnPropertyEnumerableAttributeClassMap : ExcelClassMap<MissingColumnPropertyEnumerableAttributeClass>
    {
        public CustomMissingColumnPropertyEnumerableAttributeClassMap()
        {
            Map(p => p.NoSuchColumn)
                .MakeOptional();
        }
    }

#pragma warning disable CS0649
    private class MissingColumnFieldClass
    {
        public int NoSuchColumn = 10;
    }
#pragma warning restore CS0649

    private class DefaultMissingColumnFieldClassMap : ExcelClassMap<MissingColumnFieldClass>
    {
        public DefaultMissingColumnFieldClassMap()
        {
            Map(p => p.NoSuchColumn);
        }
    }

    private class CustomMissingColumnFieldClassMap : ExcelClassMap<MissingColumnFieldClass>
    {
        public CustomMissingColumnFieldClassMap()
        {
            Map(p => p.NoSuchColumn)
                .MakeOptional();
        }
    }

    private class MissingColumnFieldAttributeClass
    {
        [ExcelOptional]
        public int NoSuchColumn = 10;
    }

    private class DefaultMissingColumnFieldAttributeClassMap : ExcelClassMap<MissingColumnFieldClass>
    {
        public DefaultMissingColumnFieldAttributeClassMap()
        {
            Map(p => p.NoSuchColumn);
        }
    }

    private class CustomMissingColumnFieldAttributeClassMap : ExcelClassMap<MissingColumnFieldClass>
    {
        public CustomMissingColumnFieldAttributeClassMap()
        {
            Map(p => p.NoSuchColumn)
                .MakeOptional();
        }
    }

#pragma warning disable CS0649
    private class MissingColumnEnumerableFieldClass
    {
        public int[] NoSuchColumn = [10];
    }
    
#pragma warning restore CS0649

    private class DefaultMissingColumnEnumerableFieldClassMap : ExcelClassMap<MissingColumnEnumerableFieldClass>
    {
        public DefaultMissingColumnEnumerableFieldClassMap()
        {
            Map(p => p.NoSuchColumn);
        }
    }

    private class CustomMissingColumnEnumerableFieldClassMap : ExcelClassMap<MissingColumnEnumerableFieldClass>
    {
        public CustomMissingColumnEnumerableFieldClassMap()
        {
            Map(p => p.NoSuchColumn)
                .MakeOptional();
        }
    }
    
    private class MissingColumnEnumerableFieldAttributeClass
    {
        [ExcelOptional]
        public int[] NoSuchColumn = [10];
    }
    private class DefaultMissingColumnEnumerableFieldAttributeClassMap : ExcelClassMap<MissingColumnEnumerableFieldAttributeClass>
    {
        public DefaultMissingColumnEnumerableFieldAttributeClassMap()
        {
            Map(p => p.NoSuchColumn);
        }
    }

    private class CustomMissingColumnEnumerableFieldAttributeClassMap : ExcelClassMap<MissingColumnEnumerableFieldAttributeClass>
    {
        public CustomMissingColumnEnumerableFieldAttributeClassMap()
        {
            Map(p => p.NoSuchColumn)
                .MakeOptional();
        }
    }

    [Fact]
    public void ReadRows_IgnoredProperty_DoesNotDeserialize()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IgnoredColumnPropertyClass>();
        Assert.Equal("CustomValue", row1.StringValue);
        Assert.Equal("a", row1.MappedValue);
    }

    private class IgnoredColumnPropertyClass
    {
        [ExcelIgnore]
        public string StringValue { get; set; } = "CustomValue";

        public string MappedValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_IgnoredField_DoesNotDeserialize()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IgnoredColumnFieldClass>();
        Assert.Equal("CustomValue", row1.StringValue);
        Assert.Equal("a", row1.MappedValue);
    }

#pragma warning disable CS0649
    private class IgnoredColumnFieldClass
    {
        [ExcelIgnore]
        public string StringValue = "CustomValue";

        public string MappedValue = default!;
    }
#pragma warning restore CS0649

    [Fact]
    public void ReadRows_IgnoredMissingProperty_DoesNotDeserialize()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnIgnoredPropertyClass>();
        Assert.Equal(10, row1.NoSuchColumn);
        Assert.Equal("a", row1.MappedValue);
    }

    private class MissingColumnIgnoredPropertyClass
    {
        [ExcelIgnore]
        public int NoSuchColumn { get; set; } = 10;

        public string MappedValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_IgnoredMissingField_DoesNotDeserialize()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MissingColumnIgnoredFieldClass>();
        Assert.Equal(10, row1.NoSuchColumn);
        Assert.Equal("a", row1.MappedValue);
    }

#pragma warning disable CS0649
    private class MissingColumnIgnoredFieldClass
    {
        [ExcelIgnore]
        public int NoSuchColumn = 10;

        public string MappedValue = default!;
    }
#pragma warning restore CS0649


    [Fact]
    public void ReadRows_IgnoredRecursiveProperty_DoesNotDeserialize()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IgnoredRecursivePropertyClass>();
        Assert.Null(row1.StringValue);
        Assert.Equal("a", row1.MappedValue);
    }

    private class IgnoredRecursivePropertyClass
    {
        [ExcelIgnore]
        public IgnoredRecursivePropertyClass StringValue { get; set; } = default!;

        public string MappedValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_IgnoredRecursiveField_DoesNotDeserialize()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IgnoredRecursiveFieldClass>();
        Assert.Null(row1.StringValue);
        Assert.Equal("a", row1.MappedValue);
    }

#pragma warning disable CS0649
    private class IgnoredRecursiveFieldClass
    {
        [ExcelIgnore]
        public IgnoredRecursiveFieldClass StringValue = default!;

        public string MappedValue = default!;
    }
#pragma warning restore CS0649

    [Fact]
    public void ReadRows_CustomMappedArrayIndexerNoSuchColumn_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<ArrayIndexerClass>(c =>
        {
            c.Map(p => p.Values[0])
                .WithColumnName("NoSuchColumn");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ArrayIndexerClass>());
    }

    private class ArrayIndexerClass
    {
        public string[] Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedArrayIndexerNoSuchColumnOptionalValue_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<ArrayIndexerClass>(c =>
        {
            c.Map(p => p.Values[0])
                .WithColumnName("NoSuchColumn")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexerClass>();
        Assert.Equal(new object?[] { null }, row.Values);
    }
}
