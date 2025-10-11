using System;
using System.Collections.Generic;
using ExcelMapper.Factories;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class ArrayIndexerMapTests
{
    [Fact]
    public void Ctor_IEnumerableFactory()
    {
        var factory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableIndexerMapT<string>(factory);
        Assert.Same(factory, map.EnumerableFactory);
        Assert.Empty(map.Values);
    }

    [Fact]
    public void Ctor_NullEnumerableFactory_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("enumerableFactory", () => new ManyToOneEnumerableIndexerMapT<string>(null!));
    }

    [Fact]
    public void TryGetValue_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableIndexerMapT<string>(enumerableFactory);
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Empty(Assert.IsType<List<string>>(result));
    }

    [Fact]
    public void TryGetValue_InvokeCanReadMultiple_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableIndexerMapT<string>(enumerableFactory);
        map.Values.Add(0, new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add(1, new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal([null, null], Assert.IsType<List<string?>>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasHeading_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var sheet = importer.ReadSheet();
        importer.Reader.Read(); // Move to first row.
        
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableIndexerMapT<string>(enumerableFactory);
        map.Values.Add(0, new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add(1, new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal([null, null], Assert.IsType<List<string?>>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasNoHeading_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        importer.Reader.Read(); // Move to first row.
        
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableIndexerMapT<string>(enumerableFactory);
        map.Values.Add(0, new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add(1, new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal([null, null], Assert.IsType<List<string?>>(result));
    }

    [Fact]
    public void TryGetValue_NullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableIndexerMapT<string>(enumerableFactory);
        object? result = null;
        Assert.Throws<ArgumentNullException>("sheet", () => map.TryGetValue(null!, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }
}