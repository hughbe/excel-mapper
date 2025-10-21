using System;
using System.Collections.Generic;
using ExcelMapper.Factories;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class DictionaryIndexerMapTests
{
    [Fact]
    public void Ctor_IEnumerableFactory()
    {
        var factory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryIndexerMapT<string, string>(factory);
        Assert.Same(factory, map.DictionaryFactory);
        Assert.Empty(map.Values);
    }

    [Fact]
    public void Ctor_NullEnumerableFactory_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("dictionaryFactory", () => new ManyToOneDictionaryIndexerMapT<string, string>(null!));
    }

    [Fact]
    public void TryGetValue_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryIndexerMapT<string, string>(dictionaryFactory);
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Empty(Assert.IsType<Dictionary<string, string>>(result));
    }

    [Fact]
    public void TryGetValue_InvokeCanReadMultiple_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryIndexerMapT<string, string>(dictionaryFactory);
        map.Values.Add("0", new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add("1", new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal(new Dictionary<string, string?> { { "0", null }, { "1", null } }, Assert.IsType<Dictionary<string, string?>>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasHeading_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var sheet = importer.ReadSheet();
        importer.Reader.Read(); // Move to first row.
        
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryIndexerMapT<string, string>(dictionaryFactory);
        map.Values.Add("0", new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add("1", new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal(new Dictionary<string, string?> { { "0", null }, { "1", null } }, Assert.IsType<Dictionary<string, string?>>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasNoHeading_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        importer.Reader.Read(); // Move to first row.
        
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryIndexerMapT<string, string>(dictionaryFactory);
        map.Values.Add("0", new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add("1", new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal(new Dictionary<string, string?> { { "0", null }, { "1", null } }, Assert.IsType<Dictionary<string, string?>>(result));
    }

    [Fact]
    public void TryGetValue_NullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryIndexerMapT<string, string>(dictionaryFactory);
        object? result = null;
        Assert.Throws<ArgumentNullException>("sheet", () => map.TryGetValue(null!, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }
}