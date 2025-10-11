using System;
using System.Collections.Generic;
using ExcelMapper.Factories;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class MultidimensionalIndexerMapTests
{
    [Fact]
    public void Ctor_IMultidimensionalArrayFactory()
    {
        var factory = new MultidimensionalArrayFactory<string>();
        var map = new ManyToOneMultidimensionalIndexerMapT<string>(factory);
        Assert.Same(factory, map.ArrayFactory);
        Assert.Empty(map.Values);
    }

    [Fact]
    public void Ctor_NullArrayFactory_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("arrayFactory", () => new ManyToOneMultidimensionalIndexerMapT<string>(null!));
    }

    [Fact]
    public void TryGetValue_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MultidimensionalArrayFactory<string>();
        var map = new ManyToOneMultidimensionalIndexerMapT<string>(factory);
        object? result = null;
        Assert.False(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCanReadMultiple_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MultidimensionalArrayFactory<string>();
        var map = new ManyToOneMultidimensionalIndexerMapT<string>(factory);
        map.Values.Add([0, 1], new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add([2, 3], new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));

        // Should make a [3, 4] array as the max indices are [2, 3].
        Assert.Equal(new string?[,] { { null, null, null, null }, { null, null, null, null }, { null, null, null, null } }, Assert.IsType<string?[,]>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasHeading_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var sheet = importer.ReadSheet();
        importer.Reader.Read(); // Move to first row.
        
        var factory = new MultidimensionalArrayFactory<string>();
        var map = new ManyToOneMultidimensionalIndexerMapT<string>(factory);
        map.Values.Add([0, 1], new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add([2, 3], new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal(new string?[,] { { null, null, null, null }, { null, null, null, null }, { null, null, null, null } }, Assert.IsType<string?[,]>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasNoHeading_Success()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        importer.Reader.Read(); // Move to first row.
        
        var factory = new MultidimensionalArrayFactory<string>();
        var map = new ManyToOneMultidimensionalIndexerMapT<string>(factory);
        map.Values.Add([0, 1], new OneToOneMap<string>(new ColumnIndexReaderFactory(0)));
        map.Values.Add([2, 3], new OneToOneMap<string>(new ColumnIndexReaderFactory(1)));
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal(new string?[,] { { null, null, null, null }, { null, null, null, null }, { null, null, null, null } }, Assert.IsType<string?[,]>(result));
    }

    [Fact]
    public void TryGetValue_NullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        var factory = new MultidimensionalArrayFactory<string>();
        var map = new ManyToOneMultidimensionalIndexerMapT<string>(factory);
        object? result = null;
        Assert.Throws<ArgumentNullException>("sheet", () => map.TryGetValue(null!, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }
}