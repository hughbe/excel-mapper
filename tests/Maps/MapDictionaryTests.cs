using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics.CodeAnalysis;
using Xunit;

namespace ExcelMapper.Tests;

public class MapDictionaryTest
{
    [Fact]
    public void ReadRow_AutoMappedIEnumerableKeyValuePairStringObject_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IEnumerable<KeyValuePair<String, object>>>());
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IEnumerable<KeyValuePair<String, object>>>());
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IEnumerable<KeyValuePair<String, object>>>());
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubIEnumerableKeyValuePairStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SubIEnumerableKeyValuePairStringObject>());
    }

    private interface SubIEnumerableKeyValuePairStringObject : IEnumerable<KeyValuePair<string, object>>
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSubIEnumerableKeyValuePairStringObjectClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<SubIEnumerableKeyValuePairStringObjectClass>(sheet.ReadRow<SubIEnumerableKeyValuePairStringObjectClass>());
    }

    private class SubIEnumerableKeyValuePairStringObjectClass : IEnumerable<KeyValuePair<string, object>>
    {
        public IEnumerator GetEnumerator() => throw new ExcelMappingException();
        
        IEnumerator<KeyValuePair<string, object>> IEnumerable<KeyValuePair<string, object>>.GetEnumerator() => throw new ExcelMappingException();
    }

    [Fact]
    public void ReadRow_AutoMappedIEnumerableKeyValuePairStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row1.Value).Count);
        Assert.Equal("a", ((Dictionary<string, object>)row1.Value)["Column1"]);
        Assert.Equal("1", ((Dictionary<string, object>)row1.Value)["Column2"]);
        Assert.Equal("2", ((Dictionary<string, object>)row1.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row1.Value)["Column4"]);

        var row2 = sheet.ReadRow<IEnumerableKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row2.Value).Count);
        Assert.Equal("b", ((Dictionary<string, object>)row2.Value)["Column1"]);
        Assert.Equal("0", ((Dictionary<string, object>)row2.Value)["Column2"]);
        Assert.Equal("0", ((Dictionary<string, object>)row2.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row2.Value)["Column4"]);

        var row3 = sheet.ReadRow<IEnumerableKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row3.Value).Count);
        Assert.Equal("c", ((Dictionary<string, object>)row3.Value)["Column1"]);
        Assert.Equal("-2", ((Dictionary<string, object>)row3.Value)["Column2"]);
        Assert.Equal("-1", ((Dictionary<string, object>)row3.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row3.Value)["Column4"]);
    }

    private class IEnumerableKeyValuePairStringObjectClass
    {
        public IEnumerable<KeyValuePair<string, object>> Value { get; set; } = default!;
    }

    private class DefaultIEnumerableKeyValuePairStringObjectClassMap : ExcelClassMap<IEnumerableKeyValuePairStringObjectClass>
    {
        public DefaultIEnumerableKeyValuePairStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIEnumerableKeyValuePairStringObjectClassMap : ExcelClassMap<IEnumerableKeyValuePairStringObjectClass>
    {
        public CustomIEnumerableKeyValuePairStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedIEnumerableKeyValuePairStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row1.Value).Count);
        Assert.Equal(1, ((Dictionary<string, int>)row1.Value)["Column1"]);
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value)["Column2"]);
        Assert.Equal(3, ((Dictionary<string, int>)row1.Value)["Column3"]);
        Assert.Equal(4, ((Dictionary<string, int>)row1.Value)["Column4"]);

        var row2 = sheet.ReadRow<IEnumerableKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row2.Value).Count);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column1"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column2"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column3"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column4"]);

        var row3 = sheet.ReadRow<IEnumerableKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row3.Value).Count);
        Assert.Equal(-2, ((Dictionary<string, int>)row3.Value)["Column1"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column2"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column3"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column4"]);
    }

    private class IEnumerableKeyValuePairStringIntClass
    {
        public IEnumerable<KeyValuePair<string, int>> Value { get; set; } = default!;
    }

    private class DefaultIEnumerableKeyValuePairStringIntClassMap : ExcelClassMap<IEnumerableKeyValuePairStringIntClass>
    {
        public DefaultIEnumerableKeyValuePairStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIEnumerableKeyValuePairStringIntClassMap : ExcelClassMap<IEnumerableKeyValuePairStringIntClass>
    {
        public CustomIEnumerableKeyValuePairStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedICollectionKeyValuePairStringObject_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<ICollection<KeyValuePair<String, object>>>());
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<ICollection<KeyValuePair<String, object>>>());
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<ICollection<KeyValuePair<String, object>>>());
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubICollectionKeyValuePairStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SubICollectionKeyValuePairStringObject>());
    }

    private interface SubICollectionKeyValuePairStringObject : ICollection<KeyValuePair<string, object>>
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSubICollectionKeyValuePairStringObjectClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<SubICollectionKeyValuePairStringObjectClass>(sheet.ReadRow<SubICollectionKeyValuePairStringObjectClass>());
    }

    private class SubICollectionKeyValuePairStringObjectClass : ICollection<KeyValuePair<string, object>>
    {
        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new ExcelMappingException();

        public bool Remove(KeyValuePair<string, object> item) => throw new NotImplementedException();

        IEnumerator<KeyValuePair<string, object>> IEnumerable<KeyValuePair<string, object>>.GetEnumerator() => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_AutoMappedICollectionKeyValuePairStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row1.Value).Count);
        Assert.Equal("a", ((Dictionary<string, object>)row1.Value)["Column1"]);
        Assert.Equal("1", ((Dictionary<string, object>)row1.Value)["Column2"]);
        Assert.Equal("2", ((Dictionary<string, object>)row1.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row1.Value)["Column4"]);

        var row2 = sheet.ReadRow<ICollectionKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row2.Value).Count);
        Assert.Equal("b", ((Dictionary<string, object>)row2.Value)["Column1"]);
        Assert.Equal("0", ((Dictionary<string, object>)row2.Value)["Column2"]);
        Assert.Equal("0", ((Dictionary<string, object>)row2.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row2.Value)["Column4"]);

        var row3 = sheet.ReadRow<ICollectionKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row3.Value).Count);
        Assert.Equal("c", ((Dictionary<string, object>)row3.Value)["Column1"]);
        Assert.Equal("-2", ((Dictionary<string, object>)row3.Value)["Column2"]);
        Assert.Equal("-1", ((Dictionary<string, object>)row3.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row3.Value)["Column4"]);
    }

    private class ICollectionKeyValuePairStringObjectClass
    {
        public ICollection<KeyValuePair<string, object>> Value { get; set; } = default!;
    }

    private class DefaultICollectionKeyValuePairStringObjectClassMap : ExcelClassMap<ICollectionKeyValuePairStringObjectClass>
    {
        public DefaultICollectionKeyValuePairStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomICollectionKeyValuePairStringObjectClassMap : ExcelClassMap<ICollectionKeyValuePairStringObjectClass>
    {
        public CustomICollectionKeyValuePairStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedICollectionKeyValuePairStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row1.Value).Count);
        Assert.Equal(1, ((Dictionary<string, int>)row1.Value)["Column1"]);
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value)["Column2"]);
        Assert.Equal(3, ((Dictionary<string, int>)row1.Value)["Column3"]);
        Assert.Equal(4, ((Dictionary<string, int>)row1.Value)["Column4"]);

        var row2 = sheet.ReadRow<ICollectionKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row2.Value).Count);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column1"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column2"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column3"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column4"]);

        var row3 = sheet.ReadRow<ICollectionKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row3.Value).Count);
        Assert.Equal(-2, ((Dictionary<string, int>)row3.Value)["Column1"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column2"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column3"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column4"]);
    }

    private class ICollectionKeyValuePairStringIntClass
    {
        public ICollection<KeyValuePair<string, int>> Value { get; set; } = default!;
    }

    private class DefaultICollectionKeyValuePairStringIntClassMap : ExcelClassMap<ICollectionKeyValuePairStringIntClass>
    {
        public DefaultICollectionKeyValuePairStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomICollectionKeyValuePairStringIntClassMap : ExcelClassMap<ICollectionKeyValuePairStringIntClass>
    {
        public CustomICollectionKeyValuePairStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyCollectionKeyValuePairStringObject_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IReadOnlyCollection<KeyValuePair<String, object>>>());
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IReadOnlyCollection<KeyValuePair<String, object>>>());
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IReadOnlyCollection<KeyValuePair<String, object>>>());
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubIReadOnlyCollectionKeyValuePairStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SubIReadOnlyCollectionKeyValuePairStringObject>());
    }

    private interface SubIReadOnlyCollectionKeyValuePairStringObject : IReadOnlyCollection<KeyValuePair<string, object>>
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSubIReadOnlyCollectionKeyValuePairStringObjectClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<SubIReadOnlyCollectionKeyValuePairStringObjectClass>(sheet.ReadRow<SubIReadOnlyCollectionKeyValuePairStringObjectClass>());
    }

    private class SubIReadOnlyCollectionKeyValuePairStringObjectClass : IReadOnlyCollection<KeyValuePair<string, object>>
    {
        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new ExcelMappingException();

        public bool Remove(KeyValuePair<string, object> item) => throw new NotImplementedException();

        IEnumerator<KeyValuePair<string, object>> IEnumerable<KeyValuePair<string, object>>.GetEnumerator() => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyCollectionKeyValuePairStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row1.Value).Count);
        Assert.Equal("a", ((Dictionary<string, object>)row1.Value)["Column1"]);
        Assert.Equal("1", ((Dictionary<string, object>)row1.Value)["Column2"]);
        Assert.Equal("2", ((Dictionary<string, object>)row1.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row1.Value)["Column4"]);

        var row2 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row2.Value).Count);
        Assert.Equal("b", ((Dictionary<string, object>)row2.Value)["Column1"]);
        Assert.Equal("0", ((Dictionary<string, object>)row2.Value)["Column2"]);
        Assert.Equal("0", ((Dictionary<string, object>)row2.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row2.Value)["Column4"]);

        var row3 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringObjectClass>();
        Assert.Equal(4, ((Dictionary<string, object>)row3.Value).Count);
        Assert.Equal("c", ((Dictionary<string, object>)row3.Value)["Column1"]);
        Assert.Equal("-2", ((Dictionary<string, object>)row3.Value)["Column2"]);
        Assert.Equal("-1", ((Dictionary<string, object>)row3.Value)["Column3"]);
        Assert.Null(((Dictionary<string, object>)row3.Value)["Column4"]);
    }

    private class IReadOnlyCollectionKeyValuePairStringObjectClass
    {
        public IReadOnlyCollection<KeyValuePair<string, object>> Value { get; set; } = default!;
    }

    private class DefaultIReadOnlyCollectionKeyValuePairStringObjectClassMap : ExcelClassMap<IReadOnlyCollectionKeyValuePairStringObjectClass>
    {
        public DefaultIReadOnlyCollectionKeyValuePairStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIReadOnlyCollectionKeyValuePairStringObjectClassMap : ExcelClassMap<IReadOnlyCollectionKeyValuePairStringObjectClass>
    {
        public CustomIReadOnlyCollectionKeyValuePairStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyCollectionKeyValuePairStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row1.Value).Count);
        Assert.Equal(1, ((Dictionary<string, int>)row1.Value)["Column1"]);
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value)["Column2"]);
        Assert.Equal(3, ((Dictionary<string, int>)row1.Value)["Column3"]);
        Assert.Equal(4, ((Dictionary<string, int>)row1.Value)["Column4"]);

        var row2 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row2.Value).Count);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column1"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column2"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column3"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column4"]);

        var row3 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringIntClass>();
        Assert.Equal(4, ((Dictionary<string, int>)row3.Value).Count);
        Assert.Equal(-2, ((Dictionary<string, int>)row3.Value)["Column1"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column2"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column3"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column4"]);
    }

    private class IReadOnlyCollectionKeyValuePairStringIntClass
    {
        public IReadOnlyCollection<KeyValuePair<string, int>> Value { get; set; } = default!;
    }

    private class DefaultIReadOnlyCollectionKeyValuePairStringIntClassMap : ExcelClassMap<IReadOnlyCollectionKeyValuePairStringIntClass>
    {
        public DefaultIReadOnlyCollectionKeyValuePairStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIReadOnlyCollectionKeyValuePairStringIntClassMap : ExcelClassMap<IReadOnlyCollectionKeyValuePairStringIntClass>
    {
        public CustomIReadOnlyCollectionKeyValuePairStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedIListKeyValuePairStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IList<KeyValuePair<String, object>>>());
    }

    [Fact]
    public void ReadRow_AutoMappedSubIListKeyValuePairStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SubIListKeyValuePairStringObject>());
    }

    private interface SubIListKeyValuePairStringObject : IList<KeyValuePair<string, object>>
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSubIListKeyValuePairStringObjectClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<SubIListKeyValuePairStringObjectClass>(sheet.ReadRow<SubIListKeyValuePairStringObjectClass>());
    }

    private class SubIListKeyValuePairStringObjectClass : IList<KeyValuePair<string, object>>
    {
        public KeyValuePair<string, object> this[int index] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new ExcelMappingException();

        public int IndexOf(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void Insert(int index, KeyValuePair<string, object> item) => throw new NotImplementedException();

        public bool Remove(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void RemoveAt(int index) => throw new NotImplementedException();

        IEnumerator<KeyValuePair<string, object>> IEnumerable<KeyValuePair<string, object>>.GetEnumerator() => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyListKeyValuePairStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyList<KeyValuePair<String, object>>>());
    }

    [Fact]
    public void ReadRow_AutoMappedSubIReadOnlyListKeyValuePairStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SubIReadOnlyListKeyValuePairStringObject>());
    }

    private interface SubIReadOnlyListKeyValuePairStringObject : IReadOnlyList<KeyValuePair<string, object>>
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSubIReadOnlyListKeyValuePairStringObjectClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<SubIReadOnlyListKeyValuePairStringObjectClass>(sheet.ReadRow<SubIReadOnlyListKeyValuePairStringObjectClass>());
    }

    private class SubIReadOnlyListKeyValuePairStringObjectClass : IReadOnlyList<KeyValuePair<string, object>>
    {
        public KeyValuePair<string, object> this[int index] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new ExcelMappingException();

        public int IndexOf(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void Insert(int index, KeyValuePair<string, object> item) => throw new NotImplementedException();

        public bool Remove(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void RemoveAt(int index) => throw new NotImplementedException();

        IEnumerator<KeyValuePair<string, object>> IEnumerable<KeyValuePair<string, object>>.GetEnumerator() => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_AutoMappedIDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IDictionary>());
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IDictionary>());
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IDictionary>());
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubIDictionary_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SubIDictionary>());
    }

    private interface SubIDictionary : IDictionary
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSubIDictionaryClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = Assert.IsType<SubIDictionaryClass>(sheet.ReadRow<SubIDictionaryClass>());      Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = Assert.IsType<SubIDictionaryClass>(sheet.ReadRow<SubIDictionaryClass>());
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = Assert.IsType<SubIDictionaryClass>(sheet.ReadRow<SubIDictionaryClass>());
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class SubIDictionaryClass : IDictionary
    {
        public Dictionary<string, object> Value { get; } = new();

        public object? this[object key] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public bool IsFixedSize => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public ICollection Keys => throw new NotImplementedException();

        public ICollection Values => throw new NotImplementedException();

        public int Count => throw new NotImplementedException();

        public bool IsSynchronized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public void Add(object key, object? value) => Value.Add((string)key, value!);

        public void Clear() => throw new NotImplementedException();

        public bool Contains(object key) => throw new NotImplementedException();

        public void CopyTo(Array array, int index) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new ExcelMappingException();

        public void Remove(object key) => throw new NotImplementedException();

        IDictionaryEnumerator IDictionary.GetEnumerator() => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_AutoMappedIDictionaryClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class IDictionaryClass
    {
        public IDictionary Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIDictionaryClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIDictionaryClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultIDictionaryClassMap : ExcelClassMap<IDictionaryClass>
    {
        public DefaultIDictionaryClassMap()
        {
            Map<object>(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedIDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomIDictionaryClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<IDictionaryClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomIDictionaryClassMap : ExcelClassMap<IDictionaryClass>
    {
        public CustomIDictionaryClassMap()
        {
            Map<object>(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedIDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IDictionary<string, object>>());
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IDictionary<string, object>>());
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IDictionary<string, object>>());
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubIDictionaryStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SubIDictionaryStringObject>());
    }

    private interface SubIDictionaryStringObject : IDictionary<string, object>
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSubIDictionaryStringObjectClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();


        var row1 = Assert.IsType<SubIDictionaryStringObjectClass>(sheet.ReadRow<SubIDictionaryStringObjectClass>());
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = Assert.IsType<SubIDictionaryStringObjectClass>(sheet.ReadRow<SubIDictionaryStringObjectClass>());
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = Assert.IsType<SubIDictionaryStringObjectClass>(sheet.ReadRow<SubIDictionaryStringObjectClass>());
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class SubIDictionaryStringObjectClass : IDictionary<string, object>
    {
        public Dictionary<string, object> Value { get; } = new();

        public object this[string key] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public ICollection<string> Keys => throw new NotImplementedException();

        public ICollection<object> Values => throw new NotImplementedException();

        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(string key, object value) => Value.Add(key, value);

        public void Add(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public bool ContainsKey(string key) => throw new NotImplementedException();

        public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator<KeyValuePair<string, object>> GetEnumerator() => throw new NotImplementedException();

        public bool Remove(string key) => throw new NotImplementedException();

        public bool Remove(KeyValuePair<string, object> item) => throw new NotImplementedException();

        public bool TryGetValue(string key, [MaybeNullWhen(false)] out object value) => throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_AutoMappedIDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class IDictionaryStringObjectClass
    {
        public IDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultIDictionaryStringObjectClassMap : ExcelClassMap<IDictionaryStringObjectClass>
    {
        public DefaultIDictionaryStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedIDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomIDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomIDictionaryStringObjectClassMap : ExcelClassMap<IDictionaryStringObjectClass>
    {
        public CustomIDictionaryStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedIDictionaryStringIntClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class IDictionaryStringIntClass
    {
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultIDictionaryStringIntClassMap : ExcelClassMap<IDictionaryStringIntClass>
    {
        public DefaultIDictionaryStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedIDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomIDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomIDictionaryStringIntClassMap : ExcelClassMap<IDictionaryStringIntClass>
    {
        public CustomIDictionaryStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedSubIReadOnlyDictionaryStringObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SubIReadOnlyDictionaryStringObject>());
    }

    private interface SubIReadOnlyDictionaryStringObject : IReadOnlyDictionary<string, object>
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSubIReadOnlyDictionaryStringObjectClass_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<SubIReadOnlyDictionaryStringObjectClass>(sheet.ReadRow<SubIReadOnlyDictionaryStringObjectClass>());
    }

    private class SubIReadOnlyDictionaryStringObjectClass : IReadOnlyDictionary<string, object>
    {
        public object this[string key] => throw new NotImplementedException();

        public IEnumerable<string> Keys => throw new NotImplementedException();

        public IEnumerable<object> Values => throw new NotImplementedException();

        public int Count => throw new NotImplementedException();

        public bool ContainsKey(string key) => throw new NotImplementedException();

        public IEnumerator<KeyValuePair<string, object>> GetEnumerator() => throw new NotImplementedException();

        public bool TryGetValue(string key, [MaybeNullWhen(false)] out object value) => throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class IReadOnlyDictionaryStringObjectClass
    {
        public IReadOnlyDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIReadOnlyDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIReadOnlyDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultIReadOnlyDictionaryStringObjectClassMap : ExcelClassMap<IReadOnlyDictionaryStringObjectClass>
    {
        public DefaultIReadOnlyDictionaryStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedIReadOnlyDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomIReadOnlyDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomIReadOnlyDictionaryStringObjectClassMap : ExcelClassMap<IReadOnlyDictionaryStringObjectClass>
    {
        public CustomIReadOnlyDictionaryStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class IReadOnlyDictionaryStringIntClass
    {
        public IReadOnlyDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIReadOnlyDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIReadOnlyDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultIReadOnlyDictionaryStringIntClassMap : ExcelClassMap<IReadOnlyDictionaryStringIntClass>
    {
        public DefaultIReadOnlyDictionaryStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedIReadOnlyDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomIReadOnlyDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomIReadOnlyDictionaryStringIntClassMap : ExcelClassMap<IReadOnlyDictionaryStringIntClass>
    {
        public CustomIReadOnlyDictionaryStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IReadOnlyDictionary<string, object>>());
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IReadOnlyDictionary<string, object>>());
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = Assert.IsType<Dictionary<string, object>>(sheet.ReadRow<IReadOnlyDictionary<string, object>>());
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedHashtable_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<Hashtable>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<Hashtable>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<Hashtable>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubHashtable_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SubHashtable>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<SubHashtable>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<SubHashtable>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    private class SubHashtable : Hashtable
    {
    }

    [Fact]
    public void ReadRow_AutoMappedHashtableClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class HashtableClass
    {
        public Hashtable Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedHashtable_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultHashtableClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultHashtableClassMap : ExcelClassMap<HashtableClass>
    {
        public DefaultHashtableClassMap()
        {
            Map<object>(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedHashtable_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomHashtableClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<HashtableClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomHashtableClassMap : ExcelClassMap<HashtableClass>
    {
        public CustomHashtableClassMap()
        {
            Map<object>(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedOrderedDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionary>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<OrderedDictionary>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<OrderedDictionary>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubOrderedDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SubOrderedDictionary>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<SubOrderedDictionary>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<SubOrderedDictionary>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    private class SubOrderedDictionary : OrderedDictionary
    {
    }

    [Fact]
    public void ReadRow_AutoMappedOrderedDictionaryClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class OrderedDictionaryClass
    {
        public OrderedDictionary Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedOrderedDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultOrderedDictionaryClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultOrderedDictionaryClassMap : ExcelClassMap<OrderedDictionaryClass>
    {
        public DefaultOrderedDictionaryClassMap()
        {
            Map<object>(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedOrderedDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomOrderedDictionaryClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<OrderedDictionaryClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomOrderedDictionaryClassMap : ExcelClassMap<OrderedDictionaryClass>
    {
        public CustomOrderedDictionaryClassMap()
        {
            Map<object>(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<Dictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<Dictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<Dictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SubDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<SubDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<SubDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    private class SubDictionary<TKey, TValue> : Dictionary<TKey, TValue> where TKey : notnull
    {
    }

    [Fact]
    public void ReadRow_AutoMappedDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DictionaryStringObjectClass
    {
        public Dictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultDictionaryStringObjectClassMap : ExcelClassMap<DictionaryStringObjectClass>
    {
        public DefaultDictionaryStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<DictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomDictionaryStringObjectClassMap : ExcelClassMap<DictionaryStringObjectClass>
    {
        public CustomDictionaryStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DictionaryStringIntClass
    {
        public Dictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultDictionaryStringIntClassMap : ExcelClassMap<DictionaryStringIntClass>
    {
        public DefaultDictionaryStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomDictionaryStringIntClassMap : ExcelClassMap<DictionaryStringIntClass>
    {
        public CustomDictionaryStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedDictionaryStringInvalidObject_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<DictionaryStringInvalidClass>(sheet.ReadRow<DictionaryStringInvalidClass>());
    }

    private class DictionaryStringInvalidClass
    {
        public Dictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedConcurrentDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<ConcurrentDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<ConcurrentDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubConcurrentDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SubConcurrentDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<SubConcurrentDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<SubConcurrentDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    private class SubConcurrentDictionary<TKey, TValue> : ConcurrentDictionary<TKey, TValue> where TKey : notnull
    {
    }

    [Fact]
    public void ReadRow_AutoMappedConcurrentDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class ConcurrentDictionaryStringObjectClass
    {
        public ConcurrentDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedConcurrentDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultConcurrentDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultConcurrentDictionaryStringObjectClassMap : ExcelClassMap<ConcurrentDictionaryStringObjectClass>
    {
        public DefaultConcurrentDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedConcurrentDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomConcurrentDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<ConcurrentDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomConcurrentDictionaryStringObjectClassMap : ExcelClassMap<ConcurrentDictionaryStringObjectClass>
    {
        public CustomConcurrentDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedConcurrentDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class ConcurrentDictionaryStringIntClass
    {
        public ConcurrentDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedConcurrentDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultConcurrentDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultConcurrentDictionaryStringIntClassMap : ExcelClassMap<ConcurrentDictionaryStringIntClass>
    {
        public DefaultConcurrentDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedConcurrentDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomConcurrentDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomConcurrentDictionaryStringIntClassMap : ExcelClassMap<ConcurrentDictionaryStringIntClass>
    {
        public CustomConcurrentDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedConcurrentDictionaryStringInvalidObject_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<ConcurrentDictionaryStringInvalidClass>(sheet.ReadRow<ConcurrentDictionaryStringInvalidClass>());
    }

    private class ConcurrentDictionaryStringInvalidClass
    {
        public ConcurrentDictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedOrderedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<OrderedDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<OrderedDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubOrderedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SubOrderedDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<SubOrderedDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<SubOrderedDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    private class SubOrderedDictionary<TKey, TValue> : OrderedDictionary<TKey, TValue> where TKey : notnull
    {
    }

    [Fact]
    public void ReadRow_AutoMappedOrderedDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class OrderedDictionaryStringObjectClass
    {
        public OrderedDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedOrderedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultOrderedDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultOrderedDictionaryStringObjectClassMap : ExcelClassMap<OrderedDictionaryStringObjectClass>
    {
        public DefaultOrderedDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedOrderedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomOrderedDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<OrderedDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomOrderedDictionaryStringObjectClassMap : ExcelClassMap<OrderedDictionaryStringObjectClass>
    {
        public CustomOrderedDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedOrderedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class OrderedDictionaryStringIntClass
    {
        public OrderedDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedOrderedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultOrderedDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultOrderedDictionaryStringIntClassMap : ExcelClassMap<OrderedDictionaryStringIntClass>
    {
        public DefaultOrderedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedOrderedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomOrderedDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<OrderedDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomOrderedDictionaryStringIntClassMap : ExcelClassMap<OrderedDictionaryStringIntClass>
    {
        public CustomOrderedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedOrderedDictionaryStringInvalidObject_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<OrderedDictionaryStringInvalidClass>(sheet.ReadRow<OrderedDictionaryStringInvalidClass>());
    }

    private class OrderedDictionaryStringInvalidClass
    {
        public OrderedDictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedSortedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<SortedDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<SortedDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedSubSortedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SubSortedDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<SubSortedDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<SubSortedDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    private class SubSortedDictionary<TKey, TValue> : SortedDictionary<TKey, TValue> where TKey : notnull
    {
    }

    [Fact]
    public void ReadRow_AutoMappedSortedDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class SortedDictionaryStringObjectClass
    {
        public SortedDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSortedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultSortedDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultSortedDictionaryStringObjectClassMap : ExcelClassMap<SortedDictionaryStringObjectClass>
    {
        public DefaultSortedDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedSortedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomSortedDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<SortedDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomSortedDictionaryStringObjectClassMap : ExcelClassMap<SortedDictionaryStringObjectClass>
    {
        public CustomSortedDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class SortedDictionaryStringIntClass
    {
        public SortedDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultSortedDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultSortedDictionaryStringIntClassMap : ExcelClassMap<SortedDictionaryStringIntClass>
    {
        public DefaultSortedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomSortedDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomSortedDictionaryStringIntClassMap : ExcelClassMap<SortedDictionaryStringIntClass>
    {
        public CustomSortedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedSortedDictionaryStringInvalidObject_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.IsType<SortedDictionaryStringInvalidClass>(sheet.ReadRow<SortedDictionaryStringInvalidClass>());
    }

    private class SortedDictionaryStringInvalidClass
    {
        public SortedDictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<ImmutableDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<ImmutableDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class ImmutableDictionaryStringObjectClass
    {
        public ImmutableDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultImmutableDictionaryStringObjectClassMap : ExcelClassMap<ImmutableDictionaryStringObjectClass>
    {
        public DefaultImmutableDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomImmutableDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<ImmutableDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomImmutableDictionaryStringObjectClassMap : ExcelClassMap<ImmutableDictionaryStringObjectClass>
    {
        public CustomImmutableDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class ImmutableDictionaryStringIntClass
    {
        public ImmutableDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultImmutableDictionaryStringIntClassMap : ExcelClassMap<ImmutableDictionaryStringIntClass>
    {
        public DefaultImmutableDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomImmutableDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomImmutableDictionaryStringIntClassMap : ExcelClassMap<ImmutableDictionaryStringIntClass>
    {
        public CustomImmutableDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableDictionaryStringInvalidObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableDictionaryStringInvalidClass>());
    }

    private class ImmutableDictionaryStringInvalidClass
    {
        public ImmutableDictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableSortedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableSortedDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class ImmutableSortedDictionaryStringObjectClass
    {
        public ImmutableSortedDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableSortedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableSortedDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultImmutableSortedDictionaryStringObjectClassMap : ExcelClassMap<ImmutableSortedDictionaryStringObjectClass>
    {
        public DefaultImmutableSortedDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableSortedDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomImmutableSortedDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomImmutableSortedDictionaryStringObjectClassMap : ExcelClassMap<ImmutableSortedDictionaryStringObjectClass>
    {
        public CustomImmutableSortedDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class ImmutableSortedDictionaryStringIntClass
    {
        public ImmutableSortedDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableSortedDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultImmutableSortedDictionaryStringIntClassMap : ExcelClassMap<ImmutableSortedDictionaryStringIntClass>
    {
        public DefaultImmutableSortedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomImmutableSortedDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomImmutableSortedDictionaryStringIntClassMap : ExcelClassMap<ImmutableSortedDictionaryStringIntClass>
    {
        public CustomImmutableSortedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableSortedDictionaryStringInvalidObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableSortedDictionaryStringInvalidClass>());
    }

    private class ImmutableSortedDictionaryStringInvalidClass
    {
        public ImmutableSortedDictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedFrozenDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<FrozenDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<FrozenDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<FrozenDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedFrozenDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class FrozenDictionaryStringObjectClass
    {
        public FrozenDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedFrozenDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultFrozenDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultFrozenDictionaryStringObjectClassMap : ExcelClassMap<FrozenDictionaryStringObjectClass>
    {
        public DefaultFrozenDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedFrozenDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomFrozenDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<FrozenDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomFrozenDictionaryStringObjectClassMap : ExcelClassMap<FrozenDictionaryStringObjectClass>
    {
        public CustomFrozenDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedFrozenDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class FrozenDictionaryStringIntClass
    {
        public FrozenDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedFrozenDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultFrozenDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultFrozenDictionaryStringIntClassMap : ExcelClassMap<FrozenDictionaryStringIntClass>
    {
        public DefaultFrozenDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedFrozenDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomFrozenDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<FrozenDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomFrozenDictionaryStringIntClassMap : ExcelClassMap<FrozenDictionaryStringIntClass>
    {
        public CustomFrozenDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedFrozenDictionaryStringInvalidObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FrozenDictionaryStringInvalidClass>());
    }

    private class FrozenDictionaryStringInvalidClass
    {
        public FrozenDictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_AutoMappedReadOnlyDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ReadOnlyDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<ReadOnlyDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<ReadOnlyDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedReadOnlyDictionaryStringObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class ReadOnlyDictionaryStringObjectClass
    {
        public ReadOnlyDictionary<string, object> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedReadOnlyDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultReadOnlyDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal("a", row1.Value["Column1"]);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);
        Assert.Null(row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal("b", row2.Value["Column1"]);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);
        Assert.Null(row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal("c", row3.Value["Column1"]);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
        Assert.Null(row3.Value["Column4"]);
    }

    private class DefaultReadOnlyDictionaryStringObjectClassMap : ExcelClassMap<ReadOnlyDictionaryStringObjectClass>
    {
        public DefaultReadOnlyDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedReadOnlyDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomReadOnlyDictionaryStringObjectClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal("1", row1.Value["Column2"]);
        Assert.Equal("2", row1.Value["Column3"]);

        var row2 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal("0", row2.Value["Column2"]);
        Assert.Equal("0", row2.Value["Column3"]);

        var row3 = sheet.ReadRow<ReadOnlyDictionaryStringObjectClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal("-2", row3.Value["Column2"]);
        Assert.Equal("-1", row3.Value["Column3"]);
    }

    private class CustomReadOnlyDictionaryStringObjectClassMap : ExcelClassMap<ReadOnlyDictionaryStringObjectClass>
    {
        public CustomReadOnlyDictionaryStringObjectClassMap()
        {
            Map(p => (IDictionary<string, object>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedReadOnlyDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class ReadOnlyDictionaryStringIntClass
    {
        public ReadOnlyDictionary<string, int> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedReadOnlyDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultReadOnlyDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);
        Assert.Equal(3, row1.Value["Column3"]);
        Assert.Equal(4, row1.Value["Column4"]);

        var row2 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);
        Assert.Equal(0, row2.Value["Column4"]);

        var row3 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(4, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
        Assert.Equal(-1, row3.Value["Column4"]);
    }

    private class DefaultReadOnlyDictionaryStringIntClassMap : ExcelClassMap<ReadOnlyDictionaryStringIntClass>
    {
        public DefaultReadOnlyDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    [Fact]
    public void ReadRow_CustomMappedReadOnlyDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomReadOnlyDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<ReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
    }

    private class CustomReadOnlyDictionaryStringIntClassMap : ExcelClassMap<ReadOnlyDictionaryStringIntClass>
    {
        public CustomReadOnlyDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    [Fact]
    public void ReadRow_AutoMappedReadOnlyDictionaryStringInvalidObject_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ReadOnlyDictionaryStringInvalidClass>());
    }

    private class ReadOnlyDictionaryStringInvalidClass
    {
        public ReadOnlyDictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DictionaryNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryStringObjectClass>());
    }

    [Fact]
    public void ReadRow_DictionaryNoHeadingWithCustomMap_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryStringObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryStringObjectClass>());
    }

    [Fact]
    public void ReadRow_DictionaryMissingColumn_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<MissingColumnClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IDictionaryStringObjectClass>());
    }

    private class MissingColumnClassMap : ExcelClassMap<IDictionaryStringObjectClass>
    {
        public MissingColumnClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("NoSuchColumn");
        }
    }

    [Fact]
    public void ReadRow_DictionaryMissingColumnOptional_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap<MissingColumnOptionalClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<IDictionaryStringObjectClass>();
        Assert.Null(row.Value);
    }

    private class MissingColumnOptionalClassMap : ExcelClassMap<IDictionaryStringObjectClass>
    {
        public MissingColumnOptionalClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("NoSuchColumn")
                .MakeOptional();
        }
    }
}
