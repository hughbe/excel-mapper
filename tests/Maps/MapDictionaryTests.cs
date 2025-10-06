﻿using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.Immutable;
using Xunit;

namespace ExcelMapper.Tests;

public class MapDictionaryTest
{
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

    [Fact]
    public void ReadRow_AutoMappedIEnumerableKeyValuePairStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value).Count);
        Assert.Equal(1, ((Dictionary<string, int>)row1.Value)["Column1"]);
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value)["Column2"]);

        var row2 = sheet.ReadRow<IEnumerableKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row2.Value).Count);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column1"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column2"]);

        var row3 = sheet.ReadRow<IEnumerableKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row3.Value).Count);
        Assert.Equal(-2, ((Dictionary<string, int>)row3.Value)["Column1"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column2"]);
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

    [Fact]
    public void ReadRow_AutoMappedICollectionKeyValuePairStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value).Count);
        Assert.Equal(1, ((Dictionary<string, int>)row1.Value)["Column1"]);
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value)["Column2"]);

        var row2 = sheet.ReadRow<ICollectionKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row2.Value).Count);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column1"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column2"]);

        var row3 = sheet.ReadRow<ICollectionKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row3.Value).Count);
        Assert.Equal(-2, ((Dictionary<string, int>)row3.Value)["Column1"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column2"]);
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

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyCollectionKeyValuePairStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value).Count);
        Assert.Equal(1, ((Dictionary<string, int>)row1.Value)["Column1"]);
        Assert.Equal(2, ((Dictionary<string, int>)row1.Value)["Column2"]);

        var row2 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row2.Value).Count);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column1"]);
        Assert.Equal(0, ((Dictionary<string, int>)row2.Value)["Column2"]);

        var row3 = sheet.ReadRow<IReadOnlyCollectionKeyValuePairStringIntClass>();
        Assert.Equal(2, ((Dictionary<string, int>)row3.Value).Count);
        Assert.Equal(-2, ((Dictionary<string, int>)row3.Value)["Column1"]);
        Assert.Equal(-1, ((Dictionary<string, int>)row3.Value)["Column2"]);
    }

    [Fact]
    public void ReadRow_AutoMappedIDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionary>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<IDictionary>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<IDictionary>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedIDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<IDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<IDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyDictionaryStringObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionary<string, object>>();
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

    [Fact]
    public void ReadRow_AutoMappedIDictionaryStringIntClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
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

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
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

    [Fact]
    public void ReadRow_AutoMappedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_AutoMappedDictionaryStringInvalidObject_ThrowsMissingMethodException()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<MissingMethodException>(() => sheet.ReadRow<DictionaryStringInvalidClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_AutoMappedIImmutableDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_AutoMappedConcurrentDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
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

    [Fact]
    public void ReadRow_DefaultMappedIDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<IDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
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

    [Fact]
    public void ReadRow_DefaultMappedIReadOnlyDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIReadOnlyDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<IReadOnlyDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
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

    [Fact]
    public void ReadRow_DefaultMappedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<DictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_DefaultMappedSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultSortedDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<SortedDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_DefaultMappedIImmutableDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIImmutableDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<ImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableSortedDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableSortedDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<ImmutableSortedDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
    }

    [Fact]
    public void ReadRow_DefaultMappedConcurrentDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultConcurrentDictionaryStringIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column1"]);
        Assert.Equal(2, row1.Value["Column2"]);

        var row2 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column1"]);
        Assert.Equal(0, row2.Value["Column2"]);

        var row3 = sheet.ReadRow<ConcurrentDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column1"]);
        Assert.Equal(-1, row3.Value["Column2"]);
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

    [Fact]
    public void ReadRow_CustomMappedIImmutableDictionaryStringInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
        importer.Configuration.RegisterClassMap(new CustomIImmutableDictionaryStringIntClassMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row1.Value.Count);
        Assert.Equal(1, row1.Value["Column2"]);
        Assert.Equal(2, row1.Value["Column3"]);

        var row2 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row2.Value.Count);
        Assert.Equal(0, row2.Value["Column2"]);
        Assert.Equal(0, row2.Value["Column3"]);

        var row3 = sheet.ReadRow<IImmutableDictionaryStringIntClass>();
        Assert.Equal(2, row3.Value.Count);
        Assert.Equal(-2, row3.Value["Column2"]);
        Assert.Equal(-1, row3.Value["Column3"]);
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

    [Fact]
    public void ReadRow_IDictionaryObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        IDictionary<string, object> row1 = sheet.ReadRow<IDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        IDictionary<string, object> row2 = sheet.ReadRow<IDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        IDictionary<string, object> row3 = sheet.ReadRow<IDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_IReadOnlyDictionaryObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        IReadOnlyDictionary<string, object> row1 = sheet.ReadRow<IReadOnlyDictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        IReadOnlyDictionary<string, object> row2 = sheet.ReadRow<IReadOnlyDictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        IReadOnlyDictionary<string, object> row3 = sheet.ReadRow<IReadOnlyDictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
    }

    [Fact]
    public void ReadRow_DictionaryObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Dictionary<string, object> row1 = sheet.ReadRow<Dictionary<string, object>>();
        Assert.Equal(4, row1.Count);
        Assert.Equal("a", row1["Column1"]);
        Assert.Equal("1", row1["Column2"]);
        Assert.Equal("2", row1["Column3"]);
        Assert.Null(row1["Column4"]);

        Dictionary<string, object> row2 = sheet.ReadRow<Dictionary<string, object>>();
        Assert.Equal(4, row2.Count);
        Assert.Equal("b", row2["Column1"]);
        Assert.Equal("0", row2["Column2"]);
        Assert.Equal("0", row2["Column3"]);
        Assert.Null(row2["Column4"]);

        Dictionary<string, object> row3 = sheet.ReadRow<Dictionary<string, object>>();
        Assert.Equal(4, row3.Count);
        Assert.Equal("c", row3["Column1"]);
        Assert.Equal("-2", row3["Column2"]);
        Assert.Equal("-1", row3["Column3"]);
        Assert.Null(row3["Column4"]);
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

    private class IDictionaryClass
    {
        public IDictionary<string, object> Value { get; set; } = default!;
    }

    private class DefaultIDictionaryClassMap : ExcelClassMap<IDictionaryClass>
    {
        public DefaultIDictionaryClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIDictionaryClassMap : ExcelClassMap<IDictionaryClass>
    {
        public CustomIDictionaryClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class IDictionaryStringObjectClass
    {
        public IDictionary<string, object> Value { get; set; } = default!;
    }

    private class DefaultIDictionaryStringObjectClassMap : ExcelClassMap<IDictionaryStringObjectClass>
    {
        public DefaultIDictionaryStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIDictionaryStringObjectClassMap : ExcelClassMap<IDictionaryStringObjectClass>
    {
        public CustomIDictionaryStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class IDictionaryStringIntClass
    {
        public IDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultIDictionaryStringIntClassMap : ExcelClassMap<IDictionaryStringIntClass>
    {
        public DefaultIDictionaryStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIDictionaryStringIntClassMap : ExcelClassMap<IDictionaryStringIntClass>
    {
        public CustomIDictionaryStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class IReadOnlyDictionaryStringObjectClass
    {
        public IReadOnlyDictionary<string, object> Value { get; set; } = default!;
    }

    private class DefaultIReadOnlyDictionaryStringObjectClassMap : ExcelClassMap<IReadOnlyDictionaryStringObjectClass>
    {
        public DefaultIReadOnlyDictionaryStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIReadOnlyDictionaryStringObjectClassMap : ExcelClassMap<IReadOnlyDictionaryStringObjectClass>
    {
        public CustomIReadOnlyDictionaryStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class IReadOnlyDictionaryStringIntClass
    {
        public IReadOnlyDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultIReadOnlyDictionaryStringIntClassMap : ExcelClassMap<IReadOnlyDictionaryStringIntClass>
    {
        public DefaultIReadOnlyDictionaryStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomIReadOnlyDictionaryStringIntClassMap : ExcelClassMap<IReadOnlyDictionaryStringIntClass>
    {
        public CustomIReadOnlyDictionaryStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class DictionaryStringObjectClass
    {
        public Dictionary<string, object> Value { get; set; } = default!;
    }

    private class DefaultDictionaryStringObjectClassMap : ExcelClassMap<DictionaryStringObjectClass>
    {
        public DefaultDictionaryStringObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomDictionaryStringObjectClassMap : ExcelClassMap<DictionaryStringObjectClass>
    {
        public CustomDictionaryStringObjectClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class DictionaryStringIntClass
    {
        public Dictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultDictionaryStringIntClassMap : ExcelClassMap<DictionaryStringIntClass>
    {
        public DefaultDictionaryStringIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    private class CustomDictionaryStringIntClassMap : ExcelClassMap<DictionaryStringIntClass>
    {
        public CustomDictionaryStringIntClassMap()
        {
            Map(p => p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class DictionaryStringInvalidClass
    {
        public Dictionary<string, ExcelSheet> Value { get; set; } = default!;
    }

    private class SortedDictionaryStringIntClass
    {
        public SortedDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultSortedDictionaryStringIntClassMap : ExcelClassMap<SortedDictionaryStringIntClass>
    {
        public DefaultSortedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    private class CustomSortedDictionaryStringIntClassMap : ExcelClassMap<SortedDictionaryStringIntClass>
    {
        public CustomSortedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class IImmutableDictionaryStringIntClass
    {
        public IImmutableDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultIImmutableDictionaryStringIntClassMap : ExcelClassMap<IImmutableDictionaryStringIntClass>
    {
        public DefaultIImmutableDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    private class CustomIImmutableDictionaryStringIntClassMap : ExcelClassMap<IImmutableDictionaryStringIntClass>
    {
        public CustomIImmutableDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class ImmutableDictionaryStringIntClass
    {
        public ImmutableDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultImmutableDictionaryStringIntClassMap : ExcelClassMap<ImmutableDictionaryStringIntClass>
    {
        public DefaultImmutableDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    private class CustomImmutableDictionaryStringIntClassMap : ExcelClassMap<ImmutableDictionaryStringIntClass>
    {
        public CustomImmutableDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class ImmutableSortedDictionaryStringIntClass
    {
        public ImmutableSortedDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultImmutableSortedDictionaryStringIntClassMap : ExcelClassMap<ImmutableSortedDictionaryStringIntClass>
    {
        public DefaultImmutableSortedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
    }

    private class CustomImmutableSortedDictionaryStringIntClassMap : ExcelClassMap<ImmutableSortedDictionaryStringIntClass>
    {
        public CustomImmutableSortedDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value)
                .WithColumnNames("Column2", "Column3");
        }
    }

    private class ConcurrentDictionaryStringIntClass
    {
        public ConcurrentDictionary<string, int> Value { get; set; } = default!;
    }

    private class DefaultConcurrentDictionaryStringIntClassMap : ExcelClassMap<ConcurrentDictionaryStringIntClass>
    {
        public DefaultConcurrentDictionaryStringIntClassMap()
        {
            Map(p => (IDictionary<string, int>)p.Value);
        }
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
