using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Xunit;

namespace ExcelMapper.Tests;

public class MapExpressionsTests
{
#pragma warning disable xUnit2013 // Do not use equality check to check for collection size.
    [Fact]
    public void ReadRow_DefaultMapped_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMappedClassMap>();
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<SimpleClass>();
        Assert.Equal(2, row.Value);
    }

    private class SimpleClass
    {
        public int Value { get; set; }
    }

    private class DefaultMappedClassMap : ExcelClassMap<SimpleClass>
    {
        public DefaultMappedClassMap()
        {
            Map(o => o.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMapped_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomMappedClassMap>();
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<SimpleClass>();
        Assert.Equal(2, row.Value);
    }

    private class CustomMappedClassMap : ExcelClassMap<SimpleClass>
    {
        public CustomMappedClassMap()
        {
            Map(o => o.Value)
                .WithColumnName("Column2");
        }
    }

    [Fact]
    public void ReadRow_TwiceMapped_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceMappedClassMap>();
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<SimpleClass>();
        Assert.Equal(2, row.Value);
    }

    private class TwiceMappedClassMap : ExcelClassMap<SimpleClass>
    {
        public TwiceMappedClassMap()
        {
            Map(o => o.Value)
                .WithColumnName("NoSuchColumn");
            Map(o => o.Value)
                .WithColumnName("Column2");
        }
    }

    [Fact]
    public void ReadRow_DefaultMappedMemberNestedChildClass_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NestedClassParent>(c =>
        {
            c.Map(o => o.Child.Value);
        });
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<NestedClassParent>();
        Assert.Equal(2, row.Child.Value);
    }

    private class NestedClassParent
    {
        public NestedClassChild Child { get; set; } = default!;
    }

    private class NestedClassChild
    {
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_TwiceMappedObject_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceMappedClassMap>();
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<SimpleClass>();
        Assert.Equal(2, row.Value);
    }

    private class TwiceMappedObjectClassMap : ExcelClassMap<NestedClassParent>
    {
        public TwiceMappedObjectClassMap()
        {
            Map(o => o.Child.Value)
                .WithColumnName("NoSuchColumn");
            Map(o => o.Child.Value)
                .WithColumnName("Column2");
        }
    }

    [Fact]
    public void ReadRow_DefaultMappedMultiplePropertiesNestedChildClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMultiplePropertiesNestedClassParentMap>();
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<MultiplePropertiesNestedClassParent>();
        Assert.Equal(2, row.Child.Column2);
        Assert.Equal(3, row.Child.Column3);
    }

    private class MultiplePropertiesNestedClassParent
    {
        public MultiplePropertiesNestedClassChild Child { get; set; } = default!;
    }

    private class MultiplePropertiesNestedClassChild
    {
        public int Column2 { get; set; }
        public int Column3 { get; set; }
    }

    private class DefaultMultiplePropertiesNestedClassParentMap : ExcelClassMap<MultiplePropertiesNestedClassParent>
    {
        public DefaultMultiplePropertiesNestedClassParentMap()
        {
            Map(o => o.Child.Column2);
            Map(o => o.Child.Column3);
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedNestedGrandChildClass_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NestedGrandParentMap>();
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<NestedGrandParent>();
        Assert.Equal(2, row.Parent.Child.Value);
    }

    private class NestedGrandParent
    {
        public NestedClassParent Parent { get; set; } = default!;
    }

    private class NestedGrandParentMap : ExcelClassMap<NestedGrandParent>
    {
        public NestedGrandParentMap()
        {
            Map(o => o.Parent.Child.Value);
        }
    }

    [Fact]
    public void ReadRow_MulipleMapped_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<MultipleMapValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal("2", row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal("0", row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal("-1", row3.Value);
    }

    private class MultipleMapValueMap : ExcelClassMap<ObjectClass>
    {
        public MultipleMapValueMap()
        {
            Map(o => o.Value);
            Map(o => o.Value)
                .WithColumnName("Column2");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedIntArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(1, row.Values[0]);
        Assert.Equal(2, row.Values[1]);
    }

    private class IntArrayClass
    {
        public int[] Values { get; set; } = default!;
    }

    private class DefaultIntArrayIndexClassMap : ExcelClassMap<IntArrayClass>
    {
        public DefaultIntArrayIndexClassMap()
        {
            Map(o => o.Values[0]);
            Map(o => o.Values[1]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedIntArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomIntArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(2, row.Values[0]);
        Assert.Equal(3, row.Values[1]);
    }

    private class CustomIntArrayIndexClassMap : ExcelClassMap<IntArrayClass>
    {
        public CustomIntArrayIndexClassMap()
        {
            Map(o => o.Values[0])
                .WithColumnName("Column2");
            Map(o => o.Values[1])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedIntArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceIntArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(3, row.Values[0]);
    }

    private class TwiceIntArrayIndexClassMap : ExcelClassMap<IntArrayClass>
    {
        public TwiceIntArrayIndexClassMap()
        {
            Map(o => o.Values[0])
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values[0])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_CustomMappedCastIntArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomIntCastArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(2, row.Values[0]);
        Assert.Equal(3, row.Values[1]);
    }

    private class CustomIntCastArrayIndexClassMap : ExcelClassMap<IntArrayClass>
    {
        public CustomIntCastArrayIndexClassMap()
        {
            Map(o => o.Values[(int)0])
                .WithColumnName("Column2");
            Map(o => o.Values[(int)1])
                .WithColumnName("Column3");
        }
    }

    private class ArrayValueClass
    {
        public int Value { get; set; }
    }

    private class ObjectArrayClass
    {
        public SimpleClass[] Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMemberArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(2, row.Values[0].Value);
        Assert.Equal(2, row.Values[1].Value);
    }

    private class DefaultObjectMemberArrayIndexClassMap : ExcelClassMap<ObjectArrayClass>
    {
        public DefaultObjectMemberArrayIndexClassMap()
        {
            Map(o => o.Values[0].Value);
            Map(o => o.Values[1].Value);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedObjectArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomObjectArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(2, row.Values[0].Value);
        Assert.Equal(3, row.Values[1].Value);
    }

    private class CustomObjectArrayIndexClassMap : ExcelClassMap<ObjectArrayClass>
    {
        public CustomObjectArrayIndexClassMap()
        {
            Map(o => o.Values[0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[1].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexMultipleFields_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectArrayMultipleFieldsClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectArrayMultipleFieldsClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(0, row.Values[0].Column1);
        Assert.Equal(1, row.Values[0].Column2);
        Assert.Equal(0, row.Values[1].Column1);
        Assert.Equal(1, row.Values[1].Column2);
    }

    private class DefaultObjectArrayMultipleFieldsClassMap : ExcelClassMap<ObjectArrayMultipleFieldsClass>
    {
        public DefaultObjectArrayMultipleFieldsClassMap()
        {
            Map(o => o.Values[0].Column1);
            Map(o => o.Values[0].Column2);
            Map(o => o.Values[1].Column1);
            Map(o => o.Values[1].Column2);
        }
    }

    private class ObjectArrayMultipleFieldsClass
    {
        public MultipleFieldsClass[] Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedObjectArrayIndexMultipleFields_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomObjectArrayIndexMultipleFieldsClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectArrayMultipleFieldsClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(1, row.Values[0].Column1);
        Assert.Equal(2, row.Values[0].Column2);
        Assert.Equal(3, row.Values[1].Column1);
        Assert.Equal(4, row.Values[1].Column2);
    }

    private class CustomObjectArrayIndexMultipleFieldsClassMap : ExcelClassMap<ObjectArrayMultipleFieldsClass>
    {
        public CustomObjectArrayIndexMultipleFieldsClassMap()
        {
            Map(o => o.Values[0].Column1)
                .WithColumnName("Column2");
            Map(o => o.Values[0].Column2)
                .WithColumnName("Column3");
            Map(o => o.Values[1].Column1)
                .WithColumnName("Column4");
            Map(o => o.Values[1].Column2)
                .WithColumnName("Column5");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexLargeMax_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMemberArrayIndexLargeMaxClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(4, row.Values.Length);
        Assert.Equal(2, row.Values[0].Value);
        Assert.Null(row.Values[1]);
        Assert.Null(row.Values[2]);
        Assert.Equal(3, row.Values[3].Value);
    }

    private class DefaultObjectMemberArrayIndexLargeMaxClassMap : ExcelClassMap<ObjectArrayClass>
    {
        public DefaultObjectMemberArrayIndexLargeMaxClassMap()
        {
            Map(o => o.Values[0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[3].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedObjectArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceObjectArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(3, row.Values[0].Value);
    }

    private class TwiceObjectArrayIndexClassMap : ExcelClassMap<ObjectArrayClass>
    {
        public TwiceObjectArrayIndexClassMap()
        {
            Map(o => o.Values[0].Value)
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values[0].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedIntArrayIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntArrayIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayArrayClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(2, row.Values[0].Length);
        Assert.Equal(1, row.Values[0][0]);
        Assert.Equal(2, row.Values[0][1]);
    }

    private class IntArrayArrayClass
    {
        public int[][] Values { get; set; } = default!;
    }

    private class DefaultIntArrayIndexArrayIndexClassMap : ExcelClassMap<IntArrayArrayClass>
    {
        public DefaultIntArrayIndexArrayIndexClassMap()
        {
            Map(o => o.Values[0][0]);
            Map(o => o.Values[0][1]);
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedIntArrayIndexMultidimensionalIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntArrayIndexMultidimensionalIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayMultidimensionalArrayClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(2, row.Values[0].Length);
        Assert.Equal(1, row.Values[0].GetLength(0));
        Assert.Equal(2, row.Values[0].GetLength(1));
        Assert.Equal(1, row.Values[0][0, 0]);
        Assert.Equal(1, row.Values[0][0, 1]);
    }

    private class IntArrayMultidimensionalArrayClass
    {
        public int[][,] Values { get; set; } = default!;
    }

    private class DefaultIntArrayIndexMultidimensionalIndexClassMap : ExcelClassMap<IntArrayMultidimensionalArrayClass>
    {
        public DefaultIntArrayIndexMultidimensionalIndexClassMap()
        {
            Map(o => o.Values[0][0, 0]);
            Map(o => o.Values[0][0, 1]);
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedIntArrayIndexMultidimensionalIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceIntArrayIndexMultidimensionalIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayMultidimensionalArrayClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(1, row.Values[0].GetLength(0));
        Assert.Equal(1, row.Values[0].GetLength(1));
        Assert.Equal(3, row.Values[0][0, 0]);
    }

    private class TwiceIntArrayIndexMultidimensionalIndexClassMap : ExcelClassMap<IntArrayMultidimensionalArrayClass>
    {
        public TwiceIntArrayIndexMultidimensionalIndexClassMap()
        {
            Map(o => o.Values[0][0, 0])
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values[0][0, 0])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedIntArrayIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntArrayIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayIndexListIndex>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(2, row.Values[0].Count);
        Assert.Equal(1, row.Values[0][0]);
        Assert.Equal(2, row.Values[0][1]);
    }

    private class IntArrayIndexListIndex
    {
        public List<int>[] Values { get; set; } = default!;
    }

    private class DefaultIntArrayIndexListIndexClassMap : ExcelClassMap<IntArrayIndexListIndex>
    {
        public DefaultIntArrayIndexListIndexClassMap()
        {
            Map(o => o.Values[0][0]);
            Map(o => o.Values[0][1]);
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedIntArrayIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntArrayIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayIndexDictionaryIndex>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(2, row.Values[0].Count);
        Assert.Equal(2, row.Values[0]["Column2"]);
        Assert.Equal(3, row.Values[0]["Column3"]);
    }

    private class IntArrayIndexDictionaryIndex
    {
        public Dictionary<string, int>[] Values { get; set; } = default!;
    }

    private class DefaultIntArrayIndexDictionaryIndexClassMap : ExcelClassMap<IntArrayIndexDictionaryIndex>
    {
        public DefaultIntArrayIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values[0]["Column2"]);
            Map(o => o.Values[0]["Column3"]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedIntArrayIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomIntArrayIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayIndexDictionaryIndex>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(2, row.Values[0].Count);
        Assert.Equal(2, row.Values[0]["Column1"]);
        Assert.Equal(3, row.Values[0]["Column2"]);
    }

    private class CustomIntArrayIndexDictionaryIndexClassMap : ExcelClassMap<IntArrayIndexDictionaryIndex>
    {
        public CustomIntArrayIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values[0]["Column1"])
                .WithColumnName("Column2");
            Map(o => o.Values[0]["Column2"])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexArrayIndexFirst_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMemberArrayIndexArrayIndexFirstClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(1, row.Values[0].Length);
        Assert.Equal(1, row.Values[0][0].Value);
        Assert.Equal(1, row.Values[0].Length);
        Assert.Equal(2, row.Values[1][0].Value);
    }

    private class ArrayIndexArrayIndexClass
    {
        public ArrayValueClass[][] Values { get; set; } = default!;
    }

    private class DefaultObjectMemberArrayIndexArrayIndexFirstClassMap : ExcelClassMap<ArrayIndexArrayIndexClass>
    {
        public DefaultObjectMemberArrayIndexArrayIndexFirstClassMap()
        {
            Map(o => o.Values[0][0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[1][0].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexArrayIndexSecond_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMemberArrayIndexArrayIndexSecondClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexArrayIndexClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(2, row.Values[0].Length);
        Assert.Equal(1, row.Values[0][0].Value);
        Assert.Equal(2, row.Values[0][1].Value);
    }

    private class DefaultObjectMemberArrayIndexArrayIndexSecondClassMap : ExcelClassMap<ArrayIndexArrayIndexClass>
    {
        public DefaultObjectMemberArrayIndexArrayIndexSecondClassMap()
        {
            Map(o => o.Values[0][0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[0][1].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexArrayIndexBoth_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMemberArrayIndexArrayIndexBothClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexArrayIndexClass>();
        Assert.Equal(4, row.Values.Length);
        Assert.Equal(3, row.Values[1].Length);
        Assert.Equal(1, row.Values[1][2].Value);
        Assert.Equal(5, row.Values[3].Length);
        Assert.Equal(2, row.Values[3][4].Value);
    }

    private class DefaultObjectMemberArrayIndexArrayIndexBothClassMap : ExcelClassMap<ArrayIndexArrayIndexClass>
    {
        public DefaultObjectMemberArrayIndexArrayIndexBothClassMap()
        {
            Map(o => o.Values[1][2].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[3][4].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMemberArrayIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexListIndexClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(1, row.Values[0][0].Value);
        Assert.Equal(2, row.Values[0][1].Value);
    }

    private class DefaultObjectMemberArrayIndexListIndexClassMap : ExcelClassMap<ArrayIndexListIndexClass>
    {
        public DefaultObjectMemberArrayIndexListIndexClassMap()
        {
            Map(o => o.Values[0][0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[0][1].Value)
                .WithColumnName("Column3");
        }
    }

    private class ArrayIndexListIndexClass
    {
        public List<ArrayValueClass>[] Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMemberArrayIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexDictionaryIndexClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(1, row.Values[0]["Column2"].Value);
        Assert.Equal(2, row.Values[0]["Column3"].Value);
    }

    private class DefaultObjectMemberArrayIndexDictionaryIndexClassMap : ExcelClassMap<ArrayIndexDictionaryIndexClass>
    {
        public DefaultObjectMemberArrayIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values[0]["Column2"].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[0]["Column3"].Value)
                .WithColumnName("Column3");
        }
    }

    private class ArrayIndexDictionaryIndexClass
    {
        public Dictionary<string, SimpleClass>[] Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexArrayIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMemberArrayIndexArrayIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexArrayIndexArrayIndexClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(1, row.Values[0][0][0].Value);
        Assert.Equal(2, row.Values[0][1][0].Value);
    }

    private class DefaultObjectMemberArrayIndexArrayIndexArrayIndexClassMap : ExcelClassMap<ArrayIndexArrayIndexArrayIndexClass>
    {
        public DefaultObjectMemberArrayIndexArrayIndexArrayIndexClassMap()
        {
            Map(o => o.Values[0][0][0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[0][1][0].Value)
                .WithColumnName("Column3");
        }
    }

    private class ArrayIndexArrayIndexArrayIndexClass
    {
        public ArrayValueClass[][][] Values { get; set; } = default!;
    }

    [Fact]
    public void Map_DefaultMappedIntMultidimensionalArrayValue_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntMultidimensionalArrayClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<MultidimensionalArrayClass>();
        Assert.Equal(1, row.Values.GetLength(0));
        Assert.Equal(2, row.Values.GetLength(1));
        Assert.Equal(0, row.Values[0, 0]);
        Assert.Equal(0, row.Values[0, 1]);
    }

    private class MultidimensionalArrayClass
    {
        public int[,] Values { get; set; } = default!;
    }

    private class DefaultIntMultidimensionalArrayClassMap : ExcelClassMap<MultidimensionalArrayClass>
    {
        public DefaultIntMultidimensionalArrayClassMap()
        {
            Map(p => p.Values[0, 0]);
            Map(p => p.Values[0, 1]);
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedIntMultidimensionalIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceIntMultidimensionalIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<MultidimensionalArrayClass>();
        Assert.Equal(1, row.Values.GetLength(0));
        Assert.Equal(1, row.Values.GetLength(1));
        Assert.Equal(3, row.Values[0, 0]);
    }

    private class TwiceIntMultidimensionalIndexClassMap : ExcelClassMap<MultidimensionalArrayClass>
    {
        public TwiceIntMultidimensionalIndexClassMap()
        {
            Map(o => o.Values[0, 0])
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values[0, 0])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedObjectMultidimensionalIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceObjectMultidimensionalIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectMultidimensionalArrayClass>();
        Assert.Equal(1, row.Values.GetLength(0));
        Assert.Equal(1, row.Values.GetLength(1));
        Assert.Equal(3, row.Values[0, 0].Value);
    }

    private class TwiceObjectMultidimensionalIndexClassMap : ExcelClassMap<ObjectMultidimensionalArrayClass>
    {
        public TwiceObjectMultidimensionalIndexClassMap()
        {
            Map(o => o.Values[0, 0].Value)
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values[0, 0].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void Map_CustomMappedIntMultidimensionalArrayValue_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomIntMultidimensionalArrayClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<MultidimensionalArrayClass>();
        Assert.Equal(3, row.Values.GetLength(0));
        Assert.Equal(4, row.Values.GetLength(1));
        Assert.Equal(1, row.Values[0, 0]);
        Assert.Equal(2, row.Values[0, 1]);
        Assert.Equal(3, row.Values[2, 3]);
    }

    private class CustomIntMultidimensionalArrayClassMap : ExcelClassMap<MultidimensionalArrayClass>
    {
        public CustomIntMultidimensionalArrayClassMap()
        {
            Map(p => p.Values[0, 0])
                .WithColumnName("Column2");
            Map(p => p.Values[0, 1])
                .WithColumnName("Column3");
            Map(p => p.Values[2, 3])
                .WithColumnName("Column4");

        }
    }

    [Fact]
    public void Map_DefaultMappedObjectMultidimensionalArrayValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMultidimensionalArrayClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectMultidimensionalArrayClass>();
        Assert.Equal(1, row.Values.GetLength(0));
        Assert.Equal(2, row.Values.GetLength(1));
        Assert.Equal(2, row.Values[0, 0].Value);
        Assert.Equal(2, row.Values[0, 1].Value);
    }

    private class ObjectMultidimensionalArrayClass
    {
        public ArrayValueClass[,] Values { get; set; } = default!;
    }

    private class DefaultObjectMultidimensionalArrayClassMap : ExcelClassMap<ObjectMultidimensionalArrayClass>
    {
        public DefaultObjectMultidimensionalArrayClassMap()
        {
            Map(p => p.Values[0, 0].Value);
            Map(p => p.Values[0, 1].Value);
        }
    }

    [Fact]
    public void Map_CustomMappedObjectMultidimensionalArrayValue_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomObjectMultidimensionalArrayClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectMultidimensionalArrayClass>();
        Assert.Equal(3, row.Values.GetLength(0));
        Assert.Equal(4, row.Values.GetLength(1));
        Assert.Equal(1, row.Values[0, 0].Value);
        Assert.Equal(2, row.Values[0, 1].Value);
        Assert.Equal(3, row.Values[2, 3].Value);
    }

    private class CustomObjectMultidimensionalArrayClassMap : ExcelClassMap<ObjectMultidimensionalArrayClass>
    {
        public CustomObjectMultidimensionalArrayClassMap()
        {
            Map(p => p.Values[0, 0].Value)
                .WithColumnName("Column2");
            Map(p => p.Values[0, 1].Value)
                .WithColumnName("Column3");
            Map(p => p.Values[2, 3].Value)
                .WithColumnName("Column4");
        }
    }


    [Fact]
    public void Map_DefaultMappedObjectMultidimensionalArrayValueMultipleFields_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectMultidimensionalArrayMultipleFieldsClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectMultidimensionalArrayMultipleFieldsClass>();
        Assert.Equal(1, row.Values.GetLength(0));
        Assert.Equal(2, row.Values.GetLength(1));
        Assert.Equal(0, row.Values[0, 0].Column1);
        Assert.Equal(1, row.Values[0, 0].Column2);
        Assert.Equal(0, row.Values[0, 1].Column1);
        Assert.Equal(1, row.Values[0, 1].Column2);
    }

    private class ObjectMultidimensionalArrayMultipleFieldsClass
    {
        public MultipleFieldsClass[,] Values { get; set; } = default!;
    }

    private class DefaultObjectMultidimensionalArrayMultipleFieldsClassMap : ExcelClassMap<ObjectMultidimensionalArrayMultipleFieldsClass>
    {
        public DefaultObjectMultidimensionalArrayMultipleFieldsClassMap()
        {
            Map(p => p.Values[0, 0].Column1);
            Map(p => p.Values[0, 0].Column2);
            Map(p => p.Values[0, 1].Column1);
            Map(p => p.Values[0, 1].Column2);
        }
    }

    [Fact]
    public void Map_CustomMappedObjectMultidimensionalArrayValueMultipleFields_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomObjectMultidimensionalArrayMultipleFieldsClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectMultidimensionalArrayMultipleFieldsClass>();
        Assert.Equal(3, row.Values.GetLength(0));
        Assert.Equal(4, row.Values.GetLength(1));
        Assert.Equal(1, row.Values[0, 0].Column1);
        Assert.Equal(2, row.Values[0, 0].Column2);
        Assert.Equal(3, row.Values[0, 1].Column1);
        Assert.Equal(4, row.Values[0, 1].Column2);
        Assert.Equal(5, row.Values[2, 3].Column1);
        Assert.Equal(6, row.Values[2, 3].Column2);
    }

    private class CustomObjectMultidimensionalArrayMultipleFieldsClassMap : ExcelClassMap<ObjectMultidimensionalArrayMultipleFieldsClass>
    {
        public CustomObjectMultidimensionalArrayMultipleFieldsClassMap()
        {
            Map(p => p.Values[0, 0].Column1)
                .WithColumnName("Column2");
            Map(p => p.Values[0, 0].Column2)
                .WithColumnName("Column3");
            Map(p => p.Values[0, 1].Column1)
                .WithColumnName("Column4");
            Map(p => p.Values[0, 1].Column2)
                .WithColumnName("Column5");
            Map(p => p.Values[2, 3].Column1)
                .WithColumnName("Column6");
            Map(p => p.Values[2, 3].Column2)
                .WithColumnName("Column7");
        }
    }

    private class MultidimensionalArrayMultipleFieldsClass
    {
        public MultipleFieldsClass[,] Values { get; set; } = default!;
    }

    private class MultipleFieldsClass
    {
        public int Column1 { get; set; }
        public int Column2 { get; set; }
    }

    [Fact]
    public void ReadRows_DefaultMappedIntListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntListClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values[0]);
        Assert.Equal(2, row.Values[1]);
    }

    private class IntListClass
    {
        public List<int> Values { get; set; } = default!;
    }

    private class DefaultIntListIndexClassMap : ExcelClassMap<IntListClass>
    {
        public DefaultIntListIndexClassMap()
        {
            Map(o => o.Values[0]);
            Map(o => o.Values[1]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedIntListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomIntListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntListClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(2, row.Values[0]);
        Assert.Equal(3, row.Values[1]);
    }

    private class CustomIntListIndexClassMap : ExcelClassMap<IntListClass>
    {
        public CustomIntListIndexClassMap()
        {
            Map(o => o.Values[0])
                .WithColumnName("Column2");
            Map(o => o.Values[1])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedIntListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceIntListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntListClass>();
        Assert.Equal(1, row.Values.Count);
        Assert.Equal(3, row.Values[0]);
    }

    private class TwiceIntListIndexClassMap : ExcelClassMap<IntListClass>
    {
        public TwiceIntListIndexClassMap()
        {
            Map(o => o.Values[0])
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values[0])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectListClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(2, row.Values[0].Value);
        Assert.Equal(3, row.Values[1].Value);
    }

    private class ListValueClass
    {
        public int Value { get; set; }
    }

    private class ObjectListClass
    {
        public List<ListValueClass> Values { get; set; } = default!;
    }

    private class DefaultObjectListIndexClassMap : ExcelClassMap<ObjectListClass>
    {
        public DefaultObjectListIndexClassMap()
        {
            Map(o => o.Values[0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[1].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberListIndexMultipleFields_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectListMultipleFieldsClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectListMultipleFieldsClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(0, row.Values[0].Column1);
        Assert.Equal(1, row.Values[0].Column2);
        Assert.Equal(0, row.Values[1].Column1);
        Assert.Equal(1, row.Values[1].Column2);
    }

    private class DefaultObjectListMultipleFieldsClassMap : ExcelClassMap<ObjectListMultipleFieldsClass>
    {
        public DefaultObjectListMultipleFieldsClassMap()
        {
            Map(o => o.Values[0].Column1);
            Map(o => o.Values[0].Column2);
            Map(o => o.Values[1].Column1);
            Map(o => o.Values[1].Column2);
        }
    }

    private class ObjectListMultipleFieldsClass
    {
        public List<MultipleFieldsClass> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedObjectListIndexMultipleFields_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomObjectListIndexMultipleFieldsClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectListMultipleFieldsClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values[0].Column1);
        Assert.Equal(2, row.Values[0].Column2);
        Assert.Equal(3, row.Values[1].Column1);
        Assert.Equal(4, row.Values[1].Column2);
    }

    private class CustomObjectListIndexMultipleFieldsClassMap : ExcelClassMap<ObjectListMultipleFieldsClass>
    {
        public CustomObjectListIndexMultipleFieldsClassMap()
        {
            Map(o => o.Values[0].Column1)
                .WithColumnName("Column2");
            Map(o => o.Values[0].Column2)
                .WithColumnName("Column3");
            Map(o => o.Values[1].Column1)
                .WithColumnName("Column4");
            Map(o => o.Values[1].Column2)
                .WithColumnName("Column5");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberListIndexLargeMax_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectListIndexLargeMaxClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectListClass>();
        Assert.Equal(4, row.Values.Count);
        Assert.Equal(2, row.Values[0].Value);
        Assert.Null(row.Values[1]);
        Assert.Null(row.Values[2]);
        Assert.Equal(3, row.Values[3].Value);
    }

    private class DefaultObjectListIndexLargeMaxClassMap : ExcelClassMap<ObjectListClass>
    {
        public DefaultObjectListIndexLargeMaxClassMap()
        {
            Map(o => o.Values[0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[3].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedObjectListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceObjectListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectListClass>();
        Assert.Equal(1, row.Values.Count);
        Assert.Equal(3, row.Values[0].Value);
    }

    private class TwiceObjectListIndexClassMap : ExcelClassMap<ObjectListClass>
    {
        public TwiceObjectListIndexClassMap()
        {
            Map(o => o.Values[0].Value)
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values[0].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexListIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultArrayIndexListIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexListIndexArrayIndexClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(1, row.Values[0][0][0].Value);
        Assert.Equal(2, row.Values[0][1][0].Value);
    }

    private class DefaultArrayIndexListIndexArrayIndexClassMap : ExcelClassMap<ArrayIndexListIndexArrayIndexClass>
    {
        public DefaultArrayIndexListIndexArrayIndexClassMap()
        {
            Map(o => o.Values[0][0][0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[0][1][0].Value)
                .WithColumnName("Column3");
        }
    }

    private class ArrayIndexListIndexArrayIndexClass
    {
        public List<ArrayValueClass[]>[] Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberArrayIndexArrayIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultArrayIndexArrayIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayIndexArrayIndexListIndexClass>();
        Assert.Equal(1, row.Values.Length);
        Assert.Equal(1, row.Values[0][0][0].Value);
        Assert.Equal(2, row.Values[0][1][0].Value);
    }

    private class DefaultArrayIndexArrayIndexListIndexClassMap : ExcelClassMap<ArrayIndexArrayIndexListIndexClass>
    {
        public DefaultArrayIndexArrayIndexListIndexClassMap()
        {
            Map(o => o.Values[0][0][0].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[0][1][0].Value)
                .WithColumnName("Column3");
        }
    }

    private class ArrayIndexArrayIndexListIndexClass
    {
        public List<ArrayValueClass>[][] Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedArrayListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ArrayListClass>();
        map.Map(o => o.Values[0]);
        map.Map(o => o.Values[1]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values[0]);
        Assert.Equal("2", row.Values[1]);
    }

    private class ArrayListClass
    {
        public ArrayList Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedArrayListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ArrayListClass>();
        map.Map(o => o.Values[0]).WithColumnName("Column2");
        map.Map(o => o.Values[1]).WithColumnName("Column3");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("2", row.Values[0]);
        Assert.Equal("3", row.Values[1]);
    }

    [Fact]
    public void ReadRows_DefaultMappedCollectionIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<CollectionClass>();
        map.Map(o => o.Values[0]);
        map.Map(o => o.Values[1]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<CollectionClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values[0]);
        Assert.Equal("2", row.Values[1]);
    }

    private class CollectionClass
    {
        public Collection<string> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedCollectionIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<CollectionClass>();
        map.Map(o => o.Values[0]).WithColumnName("Column2");
        map.Map(o => o.Values[1]).WithColumnName("Column3");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<CollectionClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("2", row.Values[0]);
        Assert.Equal("3", row.Values[1]);
    }

    [Fact]
    public void ReadRows_DefaultMappedReadOnlyCollectionIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ReadOnlyCollectionClass>();
        map.Map(o => o.Values[0]);
        map.Map(o => o.Values[1]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ReadOnlyCollectionClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values[0]);
        Assert.Equal("2", row.Values[1]);
    }

    private class ReadOnlyCollectionClass
    {
        public ReadOnlyCollection<string> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedReadOnlyCollectionIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ReadOnlyCollectionClass>();
        map.Map(o => o.Values[0]).WithColumnName("Column2");
        map.Map(o => o.Values[1]).WithColumnName("Column3");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ReadOnlyCollectionClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("2", row.Values[0]);
        Assert.Equal("3", row.Values[1]);
    }

    [Fact]
    public void ReadRows_DefaultMappedIntImmutableArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<IntImmutableArrayClass>();
        map.Map(o => o.Values[0]);
        map.Map(o => o.Values[1]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntImmutableArrayClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(1, row.Values[0]);
        Assert.Equal(2, row.Values[1]);
    }

    private class IntImmutableArrayClass
    {
        public ImmutableArray<int> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedIntImmutableArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<IntImmutableArrayClass>();
        map.Map(o => o.Values[0]).WithColumnName("Column2");
        map.Map(o => o.Values[1]).WithColumnName("Column3");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntImmutableArrayClass>();
        Assert.Equal(2, row.Values.Length);
        Assert.Equal(2, row.Values[0]);
        Assert.Equal(3, row.Values[1]);
    }

    [Fact]
    public void ReadRows_DefaultMappedIntImmutableListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<IntImmutableListClass>();
        map.Map(o => o.Values[0]);
        map.Map(o => o.Values[1]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntImmutableListClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values[0]);
        Assert.Equal(2, row.Values[1]);
    }

    private class IntImmutableListClass
    {
        public ImmutableList<int> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedIntImmutableListIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<IntImmutableListClass>();
        map.Map(o => o.Values[0]).WithColumnName("Column2");
        map.Map(o => o.Values[1]).WithColumnName("Column3");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntImmutableListClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(2, row.Values[0]);
        Assert.Equal(3, row.Values[1]);
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberListIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectListIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ListIndexDictionaryIndexClass>();
        Assert.Equal(1, row.Values.Count);
        Assert.Equal(1, row.Values[0]["Column2"].Value);
        Assert.Equal(2, row.Values[0]["Column3"].Value);
    }

    private class DefaultObjectListIndexDictionaryIndexClassMap : ExcelClassMap<ListIndexDictionaryIndexClass>
    {
        public DefaultObjectListIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values[0]["Column2"].Value)
                .WithColumnName("Column2");
            Map(o => o.Values[0]["Column3"].Value)
                .WithColumnName("Column3");
        }
    }

    private class ListIndexDictionaryIndexClass
    {
        public List<Dictionary<string, SimpleClass>> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["Column2"]);
        Assert.Equal(2, row.Values["Column3"]);
    }

    private class DictionaryClass
    {
        public Dictionary<string, int> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexClassMap : ExcelClassMap<DictionaryClass>
    {
        public DefaultDictionaryIndexClassMap()
        {
            Map(o => o.Values["Column2"]);
            Map(o => o.Values["Column3"]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"]);
        Assert.Equal(2, row.Values["key3"]);
    }

    private class CustomDictionaryIndexClassMap : ExcelClassMap<DictionaryClass>
    {
        public CustomDictionaryIndexClassMap()
        {
            Map(o => o.Values["key2"])
                .WithColumnName("Column2");
            Map(o => o.Values["key3"])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_EmptyStringDefaultMappedIntDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryClass>(c =>
        {
            c.Map(o => o.Values[""]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryClass>();
        Assert.Equal(1, row.Values.Count);
        Assert.Equal(1, row.Values[""]);
    }

    [Fact]
    public void ReadRows_EmptyStringCustomMappedIntDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryClass>(c =>
        {
            c.Map(o => o.Values[""])
                .WithColumnName("Column2");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryClass>();
        Assert.Equal(1, row.Values.Count);
        Assert.Equal(2, row.Values[""]);
    }

    private class TwiceIntDictionaryIndexClassMap : ExcelClassMap<DictionaryClass>
    {
        public TwiceIntDictionaryIndexClassMap()
        {
            Map(o => o.Values["key"])
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values["key"])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedIntDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceIntDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryClass>();
        Assert.Equal(1, row.Values.Count);
        Assert.Equal(3, row.Values["key"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["Column2"].Value);
        Assert.Equal(2, row.Values["Column3"].Value);
    }

    private class ObjectDictionaryClass
    {
        public Dictionary<string, SimpleClass> Values { get; set; } = default!;
    }

    private class DefaultObjectDictionaryIndexClassMap : ExcelClassMap<ObjectDictionaryClass>
    {
        public DefaultObjectDictionaryIndexClassMap()
        {
            Map(o => o.Values["Column2"].Value)
                .WithColumnName("Column2");
            Map(o => o.Values["Column3"].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_TwiceMappedObjectDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<TwiceObjectDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectDictionaryClass>();
        Assert.Equal(1, row.Values.Count);
        Assert.Equal(3, row.Values["key"].Value);
    }

    private class TwiceObjectDictionaryIndexClassMap : ExcelClassMap<ObjectDictionaryClass>
    {
        public TwiceObjectDictionaryIndexClassMap()
        {
            Map(o => o.Values["key"].Value)
                .WithColumnName("NoSuchColumn");
            Map(o => o.Values["key"].Value)
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedObjectMemberDictionaryIndexMultipleFields_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectDictionaryMultipleFieldsClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectDictionaryMultipleFieldsClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(0, row.Values["key1"].Column1);
        Assert.Equal(1, row.Values["key1"].Column2);
        Assert.Equal(0, row.Values["key2"].Column1);
        Assert.Equal(1, row.Values["key2"].Column2);
    }

    private class DefaultObjectDictionaryMultipleFieldsClassMap : ExcelClassMap<ObjectDictionaryMultipleFieldsClass>
    {
        public DefaultObjectDictionaryMultipleFieldsClassMap()
        {
            Map(o => o.Values["key1"].Column1);
            Map(o => o.Values["key1"].Column2);
            Map(o => o.Values["key2"].Column1);
            Map(o => o.Values["key2"].Column2);
        }
    }

    private class ObjectDictionaryMultipleFieldsClass
    {
        public Dictionary<string, MultipleFieldsClass> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedObjectDictionaryIndexMultipleFields_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomObjectDictionaryIndexMultipleFieldsClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectDictionaryMultipleFieldsClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Column1);
        Assert.Equal(2, row.Values["key1"].Column2);
        Assert.Equal(3, row.Values["key2"].Column1);
        Assert.Equal(4, row.Values["key2"].Column2);
    }

    private class CustomObjectDictionaryIndexMultipleFieldsClassMap : ExcelClassMap<ObjectDictionaryMultipleFieldsClass>
    {
        public CustomObjectDictionaryIndexMultipleFieldsClassMap()
        {
            Map(o => o.Values["key1"].Column1)
                .WithColumnName("Column2");
            Map(o => o.Values["key1"].Column2)
                .WithColumnName("Column3");
            Map(o => o.Values["key2"].Column1)
                .WithColumnName("Column4");
            Map(o => o.Values["key2"].Column2)
                .WithColumnName("Column5");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["Column2"].Length);
        Assert.Equal(0, row.Values["Column2"][0]);
        Assert.Equal(2, row.Values["Column3"].Length);
        Assert.Equal(1, row.Values["Column3"][1]);
    }


    private class DictionaryIndexArrayIndexClass
    {
        public Dictionary<string, int[]> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexArrayIndexClassMap : ExcelClassMap<DictionaryIndexArrayIndexClass>
    {
        public DefaultDictionaryIndexArrayIndexClassMap()
        {
            Map(o => o.Values["Column2"][0]);
            Map(o => o.Values["Column3"][1]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Length);
        Assert.Equal(1, row.Values["key2"][0]);
        Assert.Equal(2, row.Values["key3"].Length);
        Assert.Equal(2, row.Values["key3"][1]);
    }

    private class CustomDictionaryIndexArrayIndexClassMap : ExcelClassMap<DictionaryIndexArrayIndexClass>
    {
        public CustomDictionaryIndexArrayIndexClassMap()
        {
            Map(o => o.Values["key2"][0])
                .WithColumnName("Column2");
            Map(o => o.Values["key3"][1])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexMultidimensionalIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexMultidimensionalIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexMultidimensionalIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["Column2"].GetLength(0));
        Assert.Equal(2, row.Values["Column2"].GetLength(1));
        Assert.Equal(0, row.Values["Column2"][0, 1]);
        Assert.Equal(3, row.Values["Column3"].GetLength(0));
        Assert.Equal(4, row.Values["Column3"].GetLength(1));
        Assert.Equal(0, row.Values["Column3"][2, 3]);
    }


    private class DictionaryIndexMultidimensionalIndexClass
    {
        public Dictionary<string, int[,]> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexMultidimensionalIndexClassMap : ExcelClassMap<DictionaryIndexMultidimensionalIndexClass>
    {
        public DefaultDictionaryIndexMultidimensionalIndexClassMap()
        {
            Map(o => o.Values["Column2"][0, 1]);
            Map(o => o.Values["Column3"][2, 3]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexMultidimensionalIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexMultidimensionalIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexMultidimensionalIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].GetLength(0));
        Assert.Equal(2, row.Values["key2"].GetLength(1));
        Assert.Equal(1, row.Values["key2"][0, 1]);
        Assert.Equal(3, row.Values["key3"].GetLength(0));
        Assert.Equal(4, row.Values["key3"].GetLength(1));
        Assert.Equal(2, row.Values["key3"][2, 3]);
    }

    private class CustomDictionaryIndexMultidimensionalIndexClassMap : ExcelClassMap<DictionaryIndexMultidimensionalIndexClass>
    {
        public CustomDictionaryIndexMultidimensionalIndexClassMap()
        {
            Map(o => o.Values["key2"][0, 1])
                .WithColumnName("Column2");
            Map(o => o.Values["key3"][2, 3])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexListIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["Column2"].Count);
        Assert.Equal(0, row.Values["Column2"][0]);
        Assert.Equal(2, row.Values["Column3"].Count);
        Assert.Equal(1, row.Values["Column3"][1]);
    }

    private class DictionaryIndexListIndexClass
    {
        public Dictionary<string, List<int>> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexListIndexClassMap : ExcelClassMap<DictionaryIndexListIndexClass>
    {
        public DefaultDictionaryIndexListIndexClassMap()
        {
            Map(o => o.Values["Column2"][0]);
            Map(o => o.Values["Column3"][1]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexListIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(1, row.Values["key2"][0]);
        Assert.Equal(2, row.Values["key3"].Count);
        Assert.Equal(2, row.Values["key3"][1]);
    }

    private class CustomDictionaryIndexListIndexClassMap : ExcelClassMap<DictionaryIndexListIndexClass>
    {
        public CustomDictionaryIndexListIndexClassMap()
        {
            Map(o => o.Values["key2"][0])
                .WithColumnName("Column2");
            Map(o => o.Values["key3"][1])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexDictionaryIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Count);
        Assert.Equal(1, row.Values["key1"]["Column2"]);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(2, row.Values["key2"]["Column3"]);
    }

    private class DictionaryIndexDictionaryIndexClass
    {
        public Dictionary<string, Dictionary<string, int>> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexDictionaryIndexClassMap : ExcelClassMap<DictionaryIndexDictionaryIndexClass>
    {
        public DefaultDictionaryIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values["key1"]["Column2"]);
            Map(o => o.Values["key2"]["Column3"]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexDictionaryIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(1, row.Values["key2"]["key3"]);
        Assert.Equal(1, row.Values["key4"].Count);
        Assert.Equal(2, row.Values["key4"]["key5"]);
    }

    private class CustomDictionaryIndexDictionaryIndexClassMap : ExcelClassMap<DictionaryIndexDictionaryIndexClass>
    {
        public CustomDictionaryIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values["key2"]["key3"])
                .WithColumnName("Column2");
            Map(o => o.Values["key4"]["key5"])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexArrayIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexArrayIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexArrayIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Length);
        Assert.Equal(2, row.Values["key1"][0].Length);
        Assert.Equal(1, row.Values["key1"][0][1]);
        Assert.Equal(3, row.Values["key3"].Length);
        Assert.Equal(4, row.Values["key3"][2].Length);
        Assert.Equal(3, row.Values["key3"][2][3]);
    }

    private class DictionaryIndexArrayIndexArrayIndexClass
    {
        public Dictionary<string, int[][]> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexArrayIndexArrayIndexClassMap : ExcelClassMap<DictionaryIndexArrayIndexArrayIndexClass>
    {
        public DefaultDictionaryIndexArrayIndexArrayIndexClassMap()
        {
            Map(o => o.Values["key1"][0][1]);
            Map(o => o.Values["key3"][2][3]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexArrayIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexArrayIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexArrayIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Length);
        Assert.Equal(2, row.Values["key2"][0].Length);
        Assert.Equal(1, row.Values["key2"][0][1]);
        Assert.Equal(3, row.Values["key5"].Length);
        Assert.Equal(4, row.Values["key5"][2].Length);
        Assert.Equal(2, row.Values["key5"][2][3]);
    }

    private class CustomDictionaryIndexArrayIndexArrayIndexClassMap : ExcelClassMap<DictionaryIndexArrayIndexArrayIndexClass>
    {
        public CustomDictionaryIndexArrayIndexArrayIndexClassMap()
        {
            Map(o => o.Values["key2"][0][1])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"][2][3])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexArrayIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexArrayIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexArrayIndexListIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Length);
        Assert.Equal(2, row.Values["key1"][0].Count);
        Assert.Equal(1, row.Values["key1"][0][1]);
        Assert.Equal(3, row.Values["key3"].Length);
        Assert.Equal(4, row.Values["key3"][2].Count);
        Assert.Equal(3, row.Values["key3"][2][3]);
    }

    private class DictionaryIndexArrayIndexListIndexClass
    {
        public Dictionary<string, List<int>[]> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexArrayIndexListIndexClassMap : ExcelClassMap<DictionaryIndexArrayIndexListIndexClass>
    {
        public DefaultDictionaryIndexArrayIndexListIndexClassMap()
        {
            Map(o => o.Values["key1"][0][1]);
            Map(o => o.Values["key3"][2][3]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexArrayIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexArrayIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexArrayIndexListIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Length);
        Assert.Equal(2, row.Values["key2"][0].Count);
        Assert.Equal(1, row.Values["key2"][0][1]);
        Assert.Equal(3, row.Values["key5"].Length);
        Assert.Equal(4, row.Values["key5"][2].Count);
        Assert.Equal(2, row.Values["key5"][2][3]);
    }

    private class CustomDictionaryIndexArrayIndexListIndexClassMap : ExcelClassMap<DictionaryIndexArrayIndexListIndexClass>
    {
        public CustomDictionaryIndexArrayIndexListIndexClassMap()
        {
            Map(o => o.Values["key2"][0][1])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"][2][3])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexArrayIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexArrayIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexArrayIndexDictionaryIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Length);
        Assert.Equal(1, row.Values["key1"][0].Count);
        Assert.Equal(1, row.Values["key1"][0]["Column2"]);
        Assert.Equal(2, row.Values["key3"].Length);
        Assert.Equal(1, row.Values["key3"][1].Count);
        Assert.Equal(2, row.Values["key3"][1]["Column3"]);
    }

    private class DictionaryIndexArrayIndexDictionaryIndexClass
    {
        public Dictionary<string, Dictionary<string, int>[]> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexArrayIndexDictionaryIndexClassMap : ExcelClassMap<DictionaryIndexArrayIndexDictionaryIndexClass>
    {
        public DefaultDictionaryIndexArrayIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values["key1"][0]["Column2"]);
            Map(o => o.Values["key3"][1]["Column3"]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexArrayIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexArrayIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexArrayIndexDictionaryIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Length);
        Assert.Equal(1, row.Values["key2"][0].Count);
        Assert.Equal(1, row.Values["key2"][0]["key4"]);
        Assert.Equal(2, row.Values["key5"].Length);
        Assert.Equal(1, row.Values["key5"][1].Count);
        Assert.Equal(2, row.Values["key5"][1]["key7"]);
    }

    private class CustomDictionaryIndexArrayIndexDictionaryIndexClassMap : ExcelClassMap<DictionaryIndexArrayIndexDictionaryIndexClass>
    {
        public CustomDictionaryIndexArrayIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values["key2"][0]["key4"])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"][1]["key7"])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexListIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexListIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexListIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Count);
        Assert.Equal(2, row.Values["key1"][0].Length);
        Assert.Equal(1, row.Values["key1"][0][1]);
        Assert.Equal(3, row.Values["key3"].Count);
        Assert.Equal(4, row.Values["key3"][2].Length);
        Assert.Equal(3, row.Values["key3"][2][3]);
    }

    private class DictionaryIndexListIndexArrayIndexClass
    {
        public Dictionary<string, List<int[]>> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexListIndexArrayIndexClassMap : ExcelClassMap<DictionaryIndexListIndexArrayIndexClass>
    {
        public DefaultDictionaryIndexListIndexArrayIndexClassMap()
        {
            Map(o => o.Values["key1"][0][1]);
            Map(o => o.Values["key3"][2][3]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexListIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexListIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexListIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(2, row.Values["key2"][0].Length);
        Assert.Equal(1, row.Values["key2"][0][1]);
        Assert.Equal(3, row.Values["key5"].Count);
        Assert.Equal(4, row.Values["key5"][2].Length);
        Assert.Equal(2, row.Values["key5"][2][3]);
    }

    private class CustomDictionaryIndexListIndexArrayIndexClassMap : ExcelClassMap<DictionaryIndexListIndexArrayIndexClass>
    {
        public CustomDictionaryIndexListIndexArrayIndexClassMap()
        {
            Map(o => o.Values["key2"][0][1])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"][2][3])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexListIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexListIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexListIndexListIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Count);
        Assert.Equal(2, row.Values["key1"][0].Count);
        Assert.Equal(1, row.Values["key1"][0][1]);
        Assert.Equal(3, row.Values["key3"].Count);
        Assert.Equal(4, row.Values["key3"][2].Count);
        Assert.Equal(3, row.Values["key3"][2][3]);
    }

    private class DictionaryIndexListIndexListIndexClass
    {
        public Dictionary<string, List<List<int>>> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexListIndexListIndexClassMap : ExcelClassMap<DictionaryIndexListIndexListIndexClass>
    {
        public DefaultDictionaryIndexListIndexListIndexClassMap()
        {
            Map(o => o.Values["key1"][0][1]);
            Map(o => o.Values["key3"][2][3]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexListIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexListIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexListIndexListIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(2, row.Values["key2"][0].Count);
        Assert.Equal(1, row.Values["key2"][0][1]);
        Assert.Equal(3, row.Values["key5"].Count);
        Assert.Equal(4, row.Values["key5"][2].Count);
        Assert.Equal(2, row.Values["key5"][2][3]);
    }

    private class CustomDictionaryIndexListIndexListIndexClassMap : ExcelClassMap<DictionaryIndexListIndexListIndexClass>
    {
        public CustomDictionaryIndexListIndexListIndexClassMap()
        {
            Map(o => o.Values["key2"][0][1])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"][2][3])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexListIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexListIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexListIndexDictionaryIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Count);
        Assert.Equal(1, row.Values["key1"][0].Count);
        Assert.Equal(1, row.Values["key1"][0]["Column2"]);
        Assert.Equal(2, row.Values["key3"].Count);
        Assert.Equal(1, row.Values["key3"][1].Count);
        Assert.Equal(2, row.Values["key3"][1]["Column3"]);
    }

    private class DictionaryIndexListIndexDictionaryIndexClass
    {
        public Dictionary<string, List<Dictionary<string, int>>> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexListIndexDictionaryIndexClassMap : ExcelClassMap<DictionaryIndexListIndexDictionaryIndexClass>
    {
        public DefaultDictionaryIndexListIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values["key1"][0]["Column2"]);
            Map(o => o.Values["key3"][1]["Column3"]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexListIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexListIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexListIndexDictionaryIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(1, row.Values["key2"][0].Count);
        Assert.Equal(1, row.Values["key2"][0]["key4"]);
        Assert.Equal(2, row.Values["key5"].Count);
        Assert.Equal(1, row.Values["key5"][1].Count);
        Assert.Equal(2, row.Values["key5"][1]["key7"]);
    }

    private class CustomDictionaryIndexListIndexDictionaryIndexClassMap : ExcelClassMap<DictionaryIndexListIndexDictionaryIndexClass>
    {
        public CustomDictionaryIndexListIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values["key2"][0]["key4"])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"][1]["key7"])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexDictionaryIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexDictionaryIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexDictionaryIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Count);
        Assert.Equal(1, row.Values["key1"]["key2"].Length);
        Assert.Equal(0, row.Values["key1"]["key2"][0]);
        Assert.Equal(1, row.Values["key3"].Count);
        Assert.Equal(2, row.Values["key3"]["key4"].Length);
        Assert.Equal(1, row.Values["key3"]["key4"][1]);
    }

    private class DictionaryIndexDictionaryIndexArrayIndexClass
    {
        public Dictionary<string, Dictionary<string, int[]>> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexDictionaryIndexArrayIndexClassMap : ExcelClassMap<DictionaryIndexDictionaryIndexArrayIndexClass>
    {
        public DefaultDictionaryIndexDictionaryIndexArrayIndexClassMap()
        {
            Map(o => o.Values["key1"]["key2"][0]);
            Map(o => o.Values["key3"]["key4"][1]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexDictionaryIndexArrayIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexDictionaryIndexArrayIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexDictionaryIndexArrayIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(1, row.Values["key2"]["key3"].Length);
        Assert.Equal(1, row.Values["key2"]["key3"][0]);
        Assert.Equal(1, row.Values["key5"].Count);
        Assert.Equal(2, row.Values["key5"]["key6"].Length);
        Assert.Equal(2, row.Values["key5"]["key6"][1]);
    }

    private class CustomDictionaryIndexDictionaryIndexArrayIndexClassMap : ExcelClassMap<DictionaryIndexDictionaryIndexArrayIndexClass>
    {
        public CustomDictionaryIndexDictionaryIndexArrayIndexClassMap()
        {
            Map(o => o.Values["key2"]["key3"][0])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"]["key6"][1])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexDictionaryIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexDictionaryIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexDictionaryIndexListIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Count);
        Assert.Equal(1, row.Values["key1"]["key2"].Count);
        Assert.Equal(0, row.Values["key1"]["key2"][0]);
        Assert.Equal(1, row.Values["key3"].Count);
        Assert.Equal(2, row.Values["key3"]["key4"].Count);
        Assert.Equal(1, row.Values["key3"]["key4"][1]);
    }


    private class DictionaryIndexDictionaryIndexListIndexClass
    {
        public Dictionary<string, Dictionary<string, List<int>>> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexDictionaryIndexListIndexClassMap : ExcelClassMap<DictionaryIndexDictionaryIndexListIndexClass>
    {
        public DefaultDictionaryIndexDictionaryIndexListIndexClassMap()
        {
            Map(o => o.Values["key1"]["key2"][0]);
            Map(o => o.Values["key3"]["key4"][1]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexDictionaryIndexListIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexDictionaryIndexListIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexDictionaryIndexListIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(1, row.Values["key2"]["key3"].Count);
        Assert.Equal(1, row.Values["key2"]["key3"][0]);
        Assert.Equal(1, row.Values["key5"].Count);
        Assert.Equal(2, row.Values["key5"]["key6"].Count);
        Assert.Equal(2, row.Values["key5"]["key6"][1]);
    }

    private class CustomDictionaryIndexDictionaryIndexListIndexClassMap : ExcelClassMap<DictionaryIndexDictionaryIndexListIndexClass>
    {
        public CustomDictionaryIndexDictionaryIndexListIndexClassMap()
        {
            Map(o => o.Values["key2"]["key3"][0])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"]["key6"][1])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedDictionaryIndexDictionaryIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryIndexDictionaryIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexDictionaryIndexDictionaryIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key1"].Count);
        Assert.Equal(1, row.Values["key1"]["key2"].Count);
        Assert.Equal(1, row.Values["key1"]["key2"]["Column2"]);
        Assert.Equal(1, row.Values["key3"].Count);
        Assert.Equal(1, row.Values["key3"]["key4"].Count);
        Assert.Equal(2, row.Values["key3"]["key4"]["Column3"]);
    }


    private class DictionaryIndexDictionaryIndexDictionaryIndexClass
    {
        public Dictionary<string, Dictionary<string, Dictionary<string, int>>> Values { get; set; } = default!;
    }

    private class DefaultDictionaryIndexDictionaryIndexDictionaryIndexClassMap : ExcelClassMap<DictionaryIndexDictionaryIndexDictionaryIndexClass>
    {
        public DefaultDictionaryIndexDictionaryIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values["key1"]["key2"]["Column2"]);
            Map(o => o.Values["key3"]["key4"]["Column3"]);
        }
    }

    [Fact]
    public void ReadRows_CustomMappedDictionaryIndexDictionaryIndexDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryIndexDictionaryIndexDictionaryIndexClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryIndexDictionaryIndexDictionaryIndexClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["key2"].Count);
        Assert.Equal(1, row.Values["key2"]["key3"].Count);
        Assert.Equal(1, row.Values["key2"]["key3"]["key4"]);
        Assert.Equal(1, row.Values["key5"].Count);
        Assert.Equal(1, row.Values["key5"]["key6"].Count);
        Assert.Equal(2, row.Values["key5"]["key6"]["key7"]);
    }

    private class CustomDictionaryIndexDictionaryIndexDictionaryIndexClassMap : ExcelClassMap<DictionaryIndexDictionaryIndexDictionaryIndexClass>
    {
        public CustomDictionaryIndexDictionaryIndexDictionaryIndexClassMap()
        {
            Map(o => o.Values["key2"]["key3"]["key4"])
                .WithColumnName("Column2");
            Map(o => o.Values["key5"]["key6"]["key7"])
                .WithColumnName("Column3");
        }
    }

    [Fact]
    public void ReadRows_DefaultMappedHashtableIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        var map = new ExcelClassMap<HashtableClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<HashtableClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values["Column2"]);
        Assert.Equal("2", row.Values["Column3"]);
    }

    private class HashtableClass
    {
        public Hashtable Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedHashtableIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<HashtableClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<HashtableClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedHybridDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        var map = new ExcelClassMap<HybridDictionaryClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<HybridDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values["Column2"]);
        Assert.Equal("2", row.Values["Column3"]);
    }

    private class HybridDictionaryClass
    {
        public HybridDictionary Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedHybridDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<HybridDictionaryClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<HybridDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedListDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        var map = new ExcelClassMap<ListDictionaryClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ListDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values["Column2"]);
        Assert.Equal("2", row.Values["Column3"]);
    }

    private class ListDictionaryClass
    {
        public ListDictionary Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedListDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ListDictionaryClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ListDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedNameValueCollectionIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        var map = new ExcelClassMap<NameValueCollectionClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<NameValueCollectionClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values["Column2"]);
        Assert.Equal("2", row.Values["Column3"]);
    }

    private class NameValueCollectionClass
    {
        public NameValueCollection Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedNameValueCollectionIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<NameValueCollectionClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<NameValueCollectionClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedStringDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        var map = new ExcelClassMap<StringDictionaryClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<StringDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values["Column2"]);
        Assert.Equal("2", row.Values["Column3"]);
    }

    private class StringDictionaryClass
    {
        public StringDictionary Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedStringDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<StringDictionaryClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<StringDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedOrderedDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        var map = new ExcelClassMap<OrderedDictionaryClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("1", row.Values["Column2"]);
        Assert.Equal("2", row.Values["Column3"]);
    }

    private class OrderedDictionaryClass
    {
        public OrderedDictionary Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedOrderedDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<OrderedDictionaryClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedReadOnlyDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ReadOnlyDictionaryClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ReadOnlyDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("2", row.Values["Column2"]);
        Assert.Equal("3", row.Values["Column3"]);
    }

    private class ReadOnlyDictionaryClass
    {
        public ReadOnlyDictionary<string, string> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedReadOnlyDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ReadOnlyDictionaryClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ReadOnlyDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }
    [Fact]
    public void ReadRows_DefaultMappedConcurrentDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ConcurrentDictionaryClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ConcurrentDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("2", row.Values["Column2"]);
        Assert.Equal("3", row.Values["Column3"]);
    }

    private class ConcurrentDictionaryClass
    {
        public ConcurrentDictionary<string, string> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedConcurrentDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<ConcurrentDictionaryClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ConcurrentDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }

    private class SortedDictionaryClass
    {
        public SortedDictionary<string, string> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedSortedDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<SortedDictionaryClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<SortedDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("3", row.Values["Column2"]);
        Assert.Equal("4", row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedIntImmutableDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        var map = new ExcelClassMap<IntImmutableDictionaryClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntImmutableDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(1, row.Values["Column2"]);
        Assert.Equal(2, row.Values["Column3"]);
    }

    private class IntImmutableDictionaryClass
    {
        public ImmutableDictionary<string, int> Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedIntImmutableDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<IntImmutableDictionaryClass>();
        map.Map(o => o.Values["Column2"]).WithColumnName("Column3");
        map.Map(o => o.Values["Column3"]).WithColumnName("Column4");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntImmutableDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal(3, row.Values["Column2"]);
        Assert.Equal(4, row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedSortedDictionaryIndex_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<SortedDictionaryClass>();
        map.Map(o => o.Values["Column2"]);
        map.Map(o => o.Values["Column3"]);
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<SortedDictionaryClass>();
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("2", row.Values["Column2"]);
        Assert.Equal("3", row.Values["Column3"]);
    }

    [Fact]
    public void ReadRows_DefaultMappedIntDictionaryKey_Success()
    {
        var map = new ExcelClassMap<IntDictionaryClass>();
        map.Map(p => p.Value[0]);
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntDictionaryClass>();
        Assert.Equal("2", row.Value[0]);
    }

    private class IntDictionaryClass
    {
        public Dictionary<int, string> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedIntDictionaryKey_Success()
    {
        var map = new ExcelClassMap<IntDictionaryClass>();
        map.Map(p => p.Value[0]).WithColumnName("Column2");
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntDictionaryClass>();
        Assert.Equal("2", row.Value[0]);
    }

    [Fact]
    public void ReadRows_DefaultMappedCustomDictionaryKey_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<CustomKeyDictionaryClass>();
        map.Map(p => p.Value[0.0]);
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomKeyDictionaryClass>());
    }

    private class CustomKeyDictionaryClass
    {
        public Dictionary<double, string> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedCustomDictionaryKey_Success()
    {
        var map = new ExcelClassMap<CustomKeyDictionaryClass>();
        map.Map(p => p.Value[0.0]).WithColumnName("Column2");
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<CustomKeyDictionaryClass>();
        Assert.Equal("2", row.Value[0.0]);
    }

    [Fact]
    public void ReadRow_ComplexChain_Success()
    {
        using var importer = Helpers.GetImporter("ExpressionsMap.xlsx");
        importer.Configuration.RegisterClassMap<ComplexChainClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ComplexChainClass1>();
        Assert.Equal(2, row.Value.Values.Length);
        Assert.Equal(0, row.Value.Values[0].Value.Values[0].Value.Values["key1"].Value.Value[0, 0].Value);
        Assert.Equal(1, row.Value.Values[0].Value.Values[0].Value.Values["key2"].Value.Value[0, 1].Value);
        Assert.Equal(2, row.Value.Values[0].Value.Values[1].Value.Values["key1"].Value.Value[1, 0].Value);
        Assert.Equal(3, row.Value.Values[0].Value.Values[1].Value.Values["key2"].Value.Value[1, 1].Value);
        Assert.Equal(4, row.Value.Values[1].Value.Values[0].Value.Values["key1"].Value.Value[0, 0].Value);
        Assert.Equal(5, row.Value.Values[1].Value.Values[0].Value.Values["key2"].Value.Value[0, 1].Value);
        Assert.Equal(6, row.Value.Values[1].Value.Values[1].Value.Values["key1"].Value.Value[1, 0].Value);
        Assert.Equal(7, row.Value.Values[1].Value.Values[1].Value.Values["key2"].Value.Value[1, 1].Value);
    }

    private class ComplexChainClass1
    {
        public ComplexChainClass2 Value { get; set; } = default!;
    }

    private class ComplexChainClass2
    {
        public ComplexChainClass3[] Values { get; set; } = default!;
    }

    private class ComplexChainClass3
    {
        public ComplexChainClass4 Value { get; set; } = default!;
    }

    private class ComplexChainClass4
    {
        public List<ComplexChainClass5> Values { get; set; } = default!;
    }

    private class ComplexChainClass5
    {
        public ComplexChainClass6 Value { get; set; } = default!;
    }

    private class ComplexChainClass6
    {
        public Dictionary<string, ComplexChainClass7> Values { get; set; } = default!;
    }

    private class ComplexChainClass7
    {
        public ComplexChainClass8 Value { get; set; } = default!;
    }

    private class ComplexChainClass8
    {
        public ComplexChainClass9[,] Value { get; set; } = default!;
    }

    private class ComplexChainClass9
    {
        public int Value { get; set; }
    }

    private class ComplexChainClassMap : ExcelClassMap<ComplexChainClass1>
    {
        public ComplexChainClassMap()
        {
            Map(o => o.Value.Values[0].Value.Values[0].Value.Values["key1"].Value.Value[0, 0].Value)
                .WithColumnName("Column1");
            Map(o => o.Value.Values[0].Value.Values[0].Value.Values["key2"].Value.Value[0, 1].Value)
                .WithColumnName("Column2");
            Map(o => o.Value.Values[0].Value.Values[1].Value.Values["key1"].Value.Value[1, 0].Value)
                .WithColumnName("Column3");
            Map(o => o.Value.Values[0].Value.Values[1].Value.Values["key2"].Value.Value[1, 1].Value)
                .WithColumnName("Column4");
            Map(o => o.Value.Values[1].Value.Values[0].Value.Values["key1"].Value.Value[0, 0].Value)
                .WithColumnName("Column5");
            Map(o => o.Value.Values[1].Value.Values[0].Value.Values["key2"].Value.Value[0, 1].Value)
                .WithColumnName("Column6");
            Map(o => o.Value.Values[1].Value.Values[1].Value.Values["key1"].Value.Value[1, 0].Value)
                .WithColumnName("Column7");
            Map(o => o.Value.Values[1].Value.Values[1].Value.Values["key2"].Value.Value[1, 1].Value)
                .WithColumnName("Column8");
        }
    }

    [Fact]
    public void ReadRow_CastExpressionValid_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, row3.Value);
    }

    private class ObjectClass
    {
        public object Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_ConvertNestedExpressionMemberResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NestedObjectClass>(c =>
        {
            c.Map(o => (int)o.Value.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NestedObjectClass>();
        Assert.Equal(2, row1.Value.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NestedObjectClass>();
        Assert.Equal(-10, row2.Value.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NestedObjectClass>();
        Assert.Equal(10, row3.Value.Value);
    }

    private class NestedObjectClass
    {
        public ObjectClass Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_ConvertNestedExpressionMemberParent_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((SimpleClass)o.Value).Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((SimpleClass)row1.Value).Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((SimpleClass)row2.Value).Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((SimpleClass)row3.Value).Value);
    }

    [Fact]
    public void ReadRow_ConvertNestedExpressionMemberParentAndResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)((ObjectClass)o.Value).Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((ObjectClass)row1.Value).Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((ObjectClass)row2.Value).Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((ObjectClass)row3.Value).Value);
    }

    [Fact]
    public void ReadRow_ArrayIndexerConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ArrayObjectClass>(c =>
        {
            c.Map(o => (int)o.Value[0])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ArrayObjectClass>();
        Assert.Equal(2, row1.Value[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ArrayObjectClass>();
        Assert.Equal(-10, row2.Value[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ArrayObjectClass>();
        Assert.Equal(10, row3.Value[0]);
    }

    private class ArrayObjectClass
    {
        public object[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_ConvertToArrayIndexer_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((int[])o.Value)[0])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((int[])row1.Value)[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((int[])row2.Value)[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((int[])row3.Value)[0]);
    }

    [Fact]
    public void ReadRow_ConvertToArrayIndexerConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)(((List<object>)o.Value)[0]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<object>)row1.Value)[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<object>)row2.Value)[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<object>)row3.Value)[0]);
    }

    [Fact]
    public void ReadRow_ConvertToArrayIndexerMember_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((SimpleClass[])o.Value)[0].Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((SimpleClass[])row1.Value)[0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((SimpleClass[])row2.Value)[0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((SimpleClass[])row3.Value)[0].Value);
    }

    [Fact]
    public void ReadRow_ConvertToArrayIndexerMemberConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)(((ObjectClass[])o.Value)[0]).Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((ObjectClass[])row1.Value)[0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((ObjectClass[])row2.Value)[0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((ObjectClass[])row3.Value)[0].Value);
    }

    [Fact]
    public void Map_ConvertToArrayIndexerIndex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<IntArrayClass>(c =>
        {
            c.Map(o => o.Values[(int)(object)0]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayClass>();
        Assert.Equal(2, row.Values[0]);
    }

    [Fact]
    public void ReadRow_MultidimensionalIndexerConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<MultidimensionalObjectClass>(c =>
        {
            c.Map(o => (int)o.Value[0, 0])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<MultidimensionalObjectClass>();
        Assert.Equal(2, row1.Value[0, 0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<MultidimensionalObjectClass>();
        Assert.Equal(-10, row2.Value[0, 0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<MultidimensionalObjectClass>();
        Assert.Equal(10, row3.Value[0, 0]);
    }

    private class MultidimensionalObjectClass
    {
        public object[,] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_ConvertToMultidimensionalIndexer_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((int[,])o.Value)[0, 0])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((int[,])row1.Value)[0, 0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((int[,])row2.Value)[0, 0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((int[,])row3.Value)[0, 0]);
    }

    [Fact]
    public void ReadRow_ConvertToMultidimensionalIndexerConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)((object[,])o.Value)[0, 0])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((object[,])row1.Value)[0, 0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((object[,])row2.Value)[0, 0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((object[,])row3.Value)[0, 0]);
    }

    [Fact]
    public void ReadRow_ConvertToMultidimensionalIndexerMember_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((SimpleClass[,])o.Value)[0, 0].Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((SimpleClass[,])row1.Value)[0, 0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((SimpleClass[,])row2.Value)[0, 0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((SimpleClass[,])row3.Value)[0, 0].Value);
    }

    [Fact]
    public void ReadRow_ConvertToMultidimensionalIndexerMemberConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)(((SimpleClass[,])o.Value)[0, 0]).Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((SimpleClass[,])row1.Value)[0, 0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((SimpleClass[,])row2.Value)[0, 0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((SimpleClass[,])row3.Value)[0, 0].Value);
    }

    [Fact]
    public void Map_ConvertToMultidimensionalIndexerIndex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<MultidimensionalArrayClass>(c =>
        {
            c.Map(o => o.Values[(int)(object)0, (int)(object)0]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<MultidimensionalArrayClass>();
        Assert.Equal(2, row.Values[0, 0]);
    }

    [Fact]
    public void ReadRow_ListIndexerConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ListObjectClass>(c =>
        {
            c.Map(o => (int)o.Value[0])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(2, row1.Value[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(-10, row2.Value[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(10, row3.Value[0]);
    }

    private class ListObjectClass
    {
        public List<object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_ConvertToListIndexer_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((List<int>)o.Value)[0])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<int>)row1.Value)[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<int>)row2.Value)[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<int>)row3.Value)[0]);
    }

    [Fact]
    public void ReadRow_ConvertToListIndexerConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)(((List<object>)o.Value)[0]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<object>)row1.Value)[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<object>)row2.Value)[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<object>)row3.Value)[0]);
    }

    [Fact]
    public void ReadRow_ConvertToListIndexerMember_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((List<SimpleClass>)o.Value)[0].Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<SimpleClass>)row1.Value)[0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<SimpleClass>)row2.Value)[0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<SimpleClass>)row3.Value)[0].Value);
    }

    [Fact]
    public void ReadRow_ConvertToListIndexerMemberConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)(((List<ObjectClass>)o.Value)[0]).Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<ObjectClass>)row1.Value)[0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<ObjectClass>)row2.Value)[0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<ObjectClass>)row3.Value)[0].Value);
    }

    [Fact]
    public void Map_ConvertToListIndexerIndex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<IntListClass>(c =>
        {
            c.Map(o => o.Values[(int)(object)0]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntListClass>();
        Assert.Equal(2, row.Values[0]);
    }
    
    [Fact]
    public void ReadRow_DictionaryIndexerConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryObjectClass>(c =>
        {
            c.Map(o => (int)o.Value["Value"])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DictionaryObjectClass>();
        Assert.Equal(2, row1.Value["Value"]);

        // Empty cell value.
        var row2 = sheet.ReadRow<DictionaryObjectClass>();
        Assert.Equal(-10, row2.Value["Value"]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DictionaryObjectClass>();
        Assert.Equal(10, row3.Value["Value"]);
    }

    private class DictionaryObjectClass
    {
        public Dictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_ConvertToDictionaryIndexer_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((Dictionary<string, int>)o.Value)["Value"])
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value)["Value"]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((Dictionary<string, int>)row2.Value)["Value"]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((Dictionary<string, int>)row3.Value)["Value"]);
    }

    [Fact]
    public void ReadRow_ConvertToDictionaryIndexerConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)(((Dictionary<string, object>)o.Value)["Value"]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((Dictionary<string, object>)row1.Value)["Value"]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((Dictionary<string, object>)row2.Value)["Value"]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((Dictionary<string, object>)row3.Value)["Value"]);
    }

    [Fact]
    public void ReadRow_ConvertToDictionaryIndexerMember_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((Dictionary<string, SimpleClass>)o.Value)["key"].Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((Dictionary<string, SimpleClass>)row1.Value)["key"].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((Dictionary<string, SimpleClass>)row2.Value)["key"].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((Dictionary<string, SimpleClass>)row3.Value)["key"].Value);
    }

    [Fact]
    public void ReadRow_ConvertToDictionaryIndexerMemberConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (int)(((Dictionary<string, ObjectClass>)o.Value)["key"]).Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((Dictionary<string, ObjectClass>)row1.Value)["key"].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((Dictionary<string, ObjectClass>)row2.Value)["key"].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((Dictionary<string, ObjectClass>)row3.Value)["key"].Value);
    }

    [Fact]
    public void Map_ConvertToDictionaryIndexerNested_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((Dictionary<string, string>)(object)o.Value)["Value"]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectClass>();
        Assert.Equal("value", ((Dictionary<string, string>)row.Value)["Value"]);
    }

    [Fact]
    public void Map_ConvertToDictionaryIndexerNestedInvalidCast_Success()
    {
        var map = new ExcelClassMap<ObjectClass>();
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => ((Dictionary<string, string>)(object)(string)(object)o.Value)["Value"]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectClass>();
        Assert.Equal("value", ((Dictionary<string, string>)row.Value)["Value"]);
    }

    [Fact]
    public void Map_ConvertToDictionaryIndexerKey_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryClass>(c =>
        {
            c.Map(o => o.Values[(string)(object)"Value"]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryClass>();
        Assert.Equal(2, row.Values["Value"]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedNestedExpressionMemberResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NestedObjectClass>(c =>
        {
            c.Map(o => checked((int)o.Value.Value))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NestedObjectClass>();
        Assert.Equal(2, row1.Value.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NestedObjectClass>();
        Assert.Equal(-10, row2.Value.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NestedObjectClass>();
        Assert.Equal(10, row3.Value.Value);
    }

    [Fact]
    public void ReadRow_ConvertCheckedNestedExpressionMemberParent_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((SimpleClass)o.Value).Value))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((SimpleClass)row1.Value).Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((SimpleClass)row2.Value).Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((SimpleClass)row3.Value).Value);
    }

    [Fact]
    public void ReadRow_ConvertCheckedNestedExpressionMemberParentAndResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)(checked((ObjectClass)o.Value).Value)))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((ObjectClass)row1.Value).Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((ObjectClass)row2.Value).Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((ObjectClass)row3.Value).Value);
    }

    [Fact]
    public void ReadRow_ArrayIndexerCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ArrayObjectClass>(c =>
        {
            c.Map(o => checked((int)o.Value[0]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ArrayObjectClass>();
        Assert.Equal(2, row1.Value[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ArrayObjectClass>();
        Assert.Equal(-10, row2.Value[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ArrayObjectClass>();
        Assert.Equal(10, row3.Value[0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToArrayIndexer_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((int[])o.Value)[0]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((int[])row1.Value)[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((int[])row2.Value)[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((int[])row3.Value)[0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToArrayIndexerCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)(checked((List<object>)o.Value)[0])))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<object>)row1.Value)[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<object>)row2.Value)[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<object>)row3.Value)[0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToArrayIndexerMember_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((SimpleClass[])o.Value)[0].Value))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((SimpleClass[])row1.Value)[0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((SimpleClass[])row2.Value)[0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((SimpleClass[])row3.Value)[0].Value);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToArrayIndexerMemberCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)(((ObjectClass[])o.Value)[0]).Value))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((ObjectClass[])row1.Value)[0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((ObjectClass[])row2.Value)[0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((ObjectClass[])row3.Value)[0].Value);
    }

    [Fact]
    public void Map_ConvertCheckedToArrayIndexerIndex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<IntArrayClass>(c =>
        {
            c.Map(o => o.Values[checked((int)(object)0)]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntArrayClass>();
        Assert.Equal(2, row.Values[0]);
    }

    [Fact]
    public void ReadRow_MultidimensionalIndexerCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<MultidimensionalObjectClass>(c =>
        {
            c.Map(o => checked((int)o.Value[0, 0]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<MultidimensionalObjectClass>();
        Assert.Equal(2, row1.Value[0, 0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<MultidimensionalObjectClass>();
        Assert.Equal(-10, row2.Value[0, 0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<MultidimensionalObjectClass>();
        Assert.Equal(10, row3.Value[0, 0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToMultidimensionalIndexer_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((int[,])o.Value)[0, 0]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((int[,])row1.Value)[0, 0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((int[,])row2.Value)[0, 0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((int[,])row3.Value)[0, 0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToMultidimensionalIndexerCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)(checked((object[,])o.Value)[0, 0])))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((object[,])row1.Value)[0, 0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((object[,])row2.Value)[0, 0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((object[,])row3.Value)[0, 0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToMultidimensionalIndexerMember_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((SimpleClass[,])o.Value)[0, 0].Value))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((SimpleClass[,])row1.Value)[0, 0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((SimpleClass[,])row2.Value)[0, 0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((SimpleClass[,])row3.Value)[0, 0].Value);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToMultidimensionalIndexerMemberCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)((checked((SimpleClass[,])o.Value)[0, 0]).Value)))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((SimpleClass[,])row1.Value)[0, 0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((SimpleClass[,])row2.Value)[0, 0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((SimpleClass[,])row3.Value)[0, 0].Value);
    }

    [Fact]
    public void Map_ConvertCheckedToMultidimensionalIndexerIndex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<MultidimensionalArrayClass>(c =>
        {
            c.Map(o => o.Values[checked((int)(object)0), checked((int)(object)0)]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<MultidimensionalArrayClass>();
        Assert.Equal(2, row.Values[0, 0]);
    }

    [Fact]
    public void ReadRow_ListIndexerCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ListObjectClass>(c =>
        {
            c.Map(o => checked((int)o.Value[0]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(2, row1.Value[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(-10, row2.Value[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(10, row3.Value[0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToListIndexer_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((List<int>)o.Value)[0]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<int>)row1.Value)[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<int>)row2.Value)[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<int>)row3.Value)[0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToListIndexerCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)(checked((List<object>)o.Value)[0])))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<object>)row1.Value)[0]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<object>)row2.Value)[0]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<object>)row3.Value)[0]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToListIndexerMember_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((List<SimpleClass>)o.Value)[0].Value))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<SimpleClass>)row1.Value)[0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<SimpleClass>)row2.Value)[0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<SimpleClass>)row3.Value)[0].Value);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToListIndexerMemberCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)(checked((List<ObjectClass>)o.Value)[0]).Value))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((List<ObjectClass>)row1.Value)[0].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((List<ObjectClass>)row2.Value)[0].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((List<ObjectClass>)row3.Value)[0].Value);
    }

    [Fact]
    public void Map_ConvertCheckedToListIndexerIndex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<IntListClass>(c =>
        {
            c.Map(o => o.Values[checked((int)(object)0)]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IntListClass>();
        Assert.Equal(2, row.Values[0]);
    }
    
    [Fact]
    public void ReadRow_DictionaryIndexerCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryObjectClass>(c =>
        {
            c.Map(o => checked((int)o.Value["Value"]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DictionaryObjectClass>();
        Assert.Equal(2, row1.Value["Value"]);

        // Empty cell value.
        var row2 = sheet.ReadRow<DictionaryObjectClass>();
        Assert.Equal(-10, row2.Value["Value"]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DictionaryObjectClass>();
        Assert.Equal(10, row3.Value["Value"]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToDictionaryIndexer_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((Dictionary<string, int>)o.Value)["Value"]))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value)["Value"]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((Dictionary<string, int>)row2.Value)["Value"]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((Dictionary<string, int>)row3.Value)["Value"]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToDictionaryIndexerCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)((checked((Dictionary<string, object>)o.Value)["Value"]))))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((Dictionary<string, object>)row1.Value)["Value"]);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((Dictionary<string, object>)row2.Value)["Value"]);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((Dictionary<string, object>)row3.Value)["Value"]);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToDictionaryIndexerMember_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((Dictionary<string, SimpleClass>)o.Value)["key"].Value))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((Dictionary<string, SimpleClass>)row1.Value)["key"].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((Dictionary<string, SimpleClass>)row2.Value)["key"].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((Dictionary<string, SimpleClass>)row3.Value)["key"].Value);
    }

    [Fact]
    public void ReadRow_ConvertCheckedToDictionaryIndexerMemberCheckedConvertResult_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => checked((int)((checked((Dictionary<string, ObjectClass>)o.Value)["key"]).Value)))
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(2, ((Dictionary<string, ObjectClass>)row1.Value)["key"].Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(-10, ((Dictionary<string, ObjectClass>)row2.Value)["key"].Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectClass>();
        Assert.Equal(10, ((Dictionary<string, ObjectClass>)row3.Value)["key"].Value);
    }

    [Fact]
    public void Map_ConvertCheckedToDictionaryIndexerNested_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((Dictionary<string, string>)(object)o.Value)["Value"]));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectClass>();
        Assert.Equal("value", ((Dictionary<string, string>)row.Value)["Value"]);
    }

    [Fact]
    public void Map_ConvertCheckedToDictionaryIndexerNestedInvalidCast_Success()
    {
        var map = new ExcelClassMap<ObjectClass>();
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ObjectClass>(c =>
        {
            c.Map(o => (checked((Dictionary<string, string>)(object)(string)(object)o.Value)["Value"]));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<ObjectClass>();
        Assert.Equal("value", ((Dictionary<string, string>)row.Value)["Value"]);
    }

    [Fact]
    public void Map_ConvertCheckedToDictionaryIndexerKey_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryClass>(c =>
        {
            c.Map(o => o.Values[checked((string)(object)"Value")]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<DictionaryClass>();
        Assert.Equal(2, row.Values["Value"]);
    }
#pragma warning restore xUnit2013 // Do not use equality check to check for collection size.
}
