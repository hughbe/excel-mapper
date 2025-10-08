using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using Xunit;

namespace ExcelMapper.Tests;

public class MapSplitEnumerableTests
{
    [Fact]
    public void ReadRow_AutoMappedObjectArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedStringArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringArrayClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<StringArrayClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<StringArrayClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<StringArrayClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<StringArrayClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedIntArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IntArrayClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntArrayClass>());

        var row3 = sheet.ReadRow<IntArrayClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IntArrayClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntArrayClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIEnumerable_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<IEnumerableClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<IEnumerableClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<IEnumerableClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IEnumerableClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedIEnumerableObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedIEnumerableInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IEnumerableIntClass>());

        var row3 = sheet.ReadRow<IEnumerableIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IEnumerableIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IEnumerableIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedICollection_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<ICollectionClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<ICollectionClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<ICollectionClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ICollectionClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedICollectionObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedICollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ICollectionIntClass>());

        var row3 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ICollectionIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIList_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IListClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<IListClass>();
        Assert.Equal(new string?[] { "1", null, "2"}, row2.Value);

        var row3 = sheet.ReadRow<IListClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<IListClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IListClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedIListObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IListObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<IListObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<IListObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<IListObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IListObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedIListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IListIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IListIntClass>());

        var row3 = sheet.ReadRow<IListIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IListIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyCollectionObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyCollectionIntClass>());

        var row3 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyCollectionIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyListObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedIReadOnlyListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyListIntClass>());

        var row3 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyListIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedArrayListClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(new ArrayList { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(new ArrayList { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(new ArrayList { "1" }, row3.Value);

        var row4 = sheet.ReadRow<ArrayListClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(new ArrayList { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedListObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<ListObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ListIntClass>());

        var row3 = sheet.ReadRow<ListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ListIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedObservableCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObservableCollectionIntClass>());

        var row3 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObservableCollectionIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<QueueIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<QueueIntClass>());

        var row3 = sheet.ReadRow<QueueIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<QueueIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<QueueIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StackIntClass>();
        Assert.Equal([3, 2, 1], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StackIntClass>());

        var row3 = sheet.ReadRow<StackIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<StackIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StackIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedSortedSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SortedSetIntClass>());

        var row3 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SortedSetIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedHashSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<HashSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HashSetIntClass>());

        var row3 = sheet.ReadRow<HashSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<HashSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HashSetIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIImmutableListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableListIntClass>());

        var row3 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableListIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIImmutableStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableStackIntClass>());

        var row3 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableStackIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIImmutableQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableQueueIntClass>());

        var row3 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableQueueIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIImmutableSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableSetIntClass>());

        var row3 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableSetIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableArrayInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableArrayIntClass>());

        var row3 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableArrayIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableListIntClass>());

        var row3 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableListIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableStackIntClass>());

        var row3 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableStackIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableQueueIntClass>());

        var row3 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableQueueIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableSortedSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableSortedSetIntClass>());

        var row3 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableSortedSetIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedImmutableHashSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableHashSetIntClass>());

        var row3 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableHashSetIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedConcurrentQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentQueueIntClass>());

        var row3 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentQueueIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedConcurrentStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Equal([3, 2, 1], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentStackIntClass>());

        var row3 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentStackIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedConcurrentBagInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentBagIntClass>());

        var row3 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentBagIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedBlockingCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlockingCollectionIntClass>());

        var row3 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlockingCollectionIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedCustomConstructorIEnumerableInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomConstructorIEnumerableIntClass>());

        var row3 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomConstructorIEnumerableIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedCustomAddIEnumerableInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomAddIEnumerableIntClass>());

        var row3 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomAddIEnumerableIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedObjectArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObjectArrayClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ObjectArrayClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedStringArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultStringArrayClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringArrayClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<StringArrayClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<StringArrayClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<StringArrayClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<StringArrayClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedIntArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntArrayClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IntArrayClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntArrayClass>());

        var row3 = sheet.ReadRow<IntArrayClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IntArrayClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntArrayClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIEnumerable_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIEnumerableClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<IEnumerableClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<IEnumerableClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<IEnumerableClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IEnumerableClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedIEnumerableObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIEnumerableObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IEnumerableObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedIEnumerableInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIEnumerableIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerableIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IEnumerableIntClass>());

        var row3 = sheet.ReadRow<IEnumerableIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IEnumerableIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IEnumerableIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedICollection_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultICollectionObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<ICollectionClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<ICollectionClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<ICollectionClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ICollectionClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedICollectionObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultICollectionObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ICollectionObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedICollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultICollectionIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ICollectionIntClass>());

        var row3 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ICollectionIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIList_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIListClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IListClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<IListClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<IListClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<IListClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IListClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedIListObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIListObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IListObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<IListObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<IListObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<IListObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IListObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedIListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IListIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IListIntClass>());

        var row3 = sheet.ReadRow<IListIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IListIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIReadOnlyCollectionObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIReadOnlyCollectionObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IReadOnlyCollectionObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedIReadOnlyCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIReadOnlyCollectionIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyCollectionIntClass>());

        var row3 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyCollectionIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIReadOnlyListObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIReadOnlyListObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Equal(["1", "2", "3"], row1.Value);

        var row2 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Equal(["1", null, "2"], row2.Value);

        var row3 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Equal(["1"], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IReadOnlyListObjectClass>();
        Assert.Equal(["Invalid"], row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedIReadOnlyListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIReadOnlyListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyListIntClass>());

        var row3 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IReadOnlyListIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedArrayListClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultArrayListClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(new ArrayList { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(new ArrayList { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(new ArrayList { "1" }, row3.Value);

        var row4 = sheet.ReadRow<ArrayListClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ArrayListClass>();
        Assert.Equal(new ArrayList { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedListObjectClass_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultListObjectClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<ListObjectClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ListObjectClass>();
        Assert.Equal(new string[] { "Invalid" }, row5.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ListIntClass>());

        var row3 = sheet.ReadRow<ListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ListIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedObservableCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultObservableCollectionIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObservableCollectionIntClass>());

        var row3 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObservableCollectionIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultQueueIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<QueueIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<QueueIntClass>());

        var row3 = sheet.ReadRow<QueueIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<QueueIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<QueueIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedSortedSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultSortedSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SortedSetIntClass>());

        var row3 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SortedSetIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedHashSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultHashSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<HashSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HashSetIntClass>());

        var row3 = sheet.ReadRow<HashSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<HashSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HashSetIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultStackIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StackIntClass>();
        Assert.Equal([3, 2, 1], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StackIntClass>());

        var row3 = sheet.ReadRow<StackIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<StackIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StackIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIImmutableListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIImmutableListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableListIntClass>());

        var row3 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableListIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIImmutableStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIImmutableStackIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableStackIntClass>());

        var row3 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableStackIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIImmutableQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIImmutableQueueIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableQueueIntClass>());

        var row3 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableQueueIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIImmutableSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIImmutableSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableSetIntClass>());

        var row3 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IImmutableSetIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableArrayInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableArrayIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableArrayIntClass>());

        var row3 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableArrayIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableListIntClass>());

        var row3 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableListIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableStackIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableStackIntClass>());

        var row3 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableStackIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableQueueIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableQueueIntClass>());

        var row3 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableQueueIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableSortedSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableSortedSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableSortedSetIntClass>());

        var row3 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableSortedSetIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedImmutableHashSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultImmutableHashSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableHashSetIntClass>());

        var row3 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableHashSetIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedConcurrentQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultConcurrentQueueIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentQueueIntClass>());

        var row3 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentQueueIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedConcurrentStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultConcurrentStackIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Equal([3, 2, 1], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentStackIntClass>());

        var row3 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentStackIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedConcurrentBagInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultConcurrentBagIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentBagIntClass>());

        var row3 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConcurrentBagIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedBlockingCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultBlockingCollectionIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlockingCollectionIntClass>());

        var row3 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlockingCollectionIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedCustomConstructorIEnumerableInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomConstructorIEnumerableIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomConstructorIEnumerableIntClass>());

        var row3 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomConstructorIEnumerableIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedCustomAddIEnumerableInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<DefaultCustomAddIEnumerableIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomAddIEnumerableIntClass>());

        var row3 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Empty(row4.Value);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomAddIEnumerableIntClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedIntArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomIntArrayClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IntArrayClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        var row2 = sheet.ReadRow<IntArrayClass>();
        Assert.Equal([1, -1, 2], row2.Value);

        var row3 = sheet.ReadRow<IntArrayClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IntArrayClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IntArrayClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedICollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomICollectionIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        var row2 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Equal([1, -1, 2], row2.Value);

        var row3 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ICollectionIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomIListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IListIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        var row2 = sheet.ReadRow<IListIntClass>();
        Assert.Equal([1, -1, 2], row2.Value);

        var row3 = sheet.ReadRow<IListIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IListIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IListIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIReadOnlyCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomIReadOnlyCollectionIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        var row2 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Equal([1, -1, 2], row2.Value);

        var row3 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IReadOnlyCollectionIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIReadOnlyListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomIReadOnlyListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        var row2 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Equal([1, -1, 2], row2.Value);

        var row3 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IReadOnlyListIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<ListIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<ListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ListIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ListIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedObservableCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomObservableCollectionIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ObservableCollectionIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomQueueIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<QueueIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        var row2 = sheet.ReadRow<QueueIntClass>();
        Assert.Equal([1, -1, 2], row2.Value);

        var row3 = sheet.ReadRow<QueueIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<QueueIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<QueueIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedSortedSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomSortedSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Equal(new int[] { -1, 1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<SortedSetIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedHashSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomHashSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<HashSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<HashSetIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<HashSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<HashSetIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<HashSetIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomStackIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StackIntClass>();
        Assert.Equal([3, 2, 1], row1.Value);

        var row2 = sheet.ReadRow<StackIntClass>();
        Assert.Equal([2, -1, 1], row2.Value);

        var row3 = sheet.ReadRow<StackIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<StackIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<StackIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIImmutableListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomIImmutableListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IImmutableListIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIImmutableStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomIImmutableStackIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        var row2 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Equal(new int[] { 2, -1, 1 }, row2.Value);

        var row3 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IImmutableStackIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIImmutableQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomIImmutableQueueIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IImmutableQueueIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIImmutableSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomIImmutableSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Equal(new int[] { -1, 1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<IImmutableSetIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableArrayInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomImmutableArrayIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ImmutableArrayIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableListInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomImmutableListIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ImmutableListIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomImmutableStackIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        var row2 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Equal(new int[] { 2, -1, 1 }, row2.Value);

        var row3 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ImmutableStackIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomImmutableQueueIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ImmutableQueueIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableSortedSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomImmutableSortedSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Equal(new int[] { -1, 1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ImmutableSortedSetIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedImmutableHashSetInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomImmutableHashSetIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Equal(new int[] { -1, 1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ImmutableHashSetIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedConcurrentQueueInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomConcurrentQueueIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        var row2 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Equal([1, -1, 2], row2.Value);

        var row3 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ConcurrentQueueIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedConcurrentStackInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomConcurrentStackIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Equal([3, 2, 1], row1.Value);

        var row2 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Equal([2, -1, 1], row2.Value);

        var row3 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ConcurrentStackIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedConcurrentBagInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomConcurrentBagIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Equal(new int[] { 3, 2, 1 }, row1.Value);

        var row2 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Equal(new int[] { 2, -1, 1 }, row2.Value);

        var row3 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<ConcurrentBagIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedBlockingCollectionInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomBlockingCollectionIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<BlockingCollectionIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedCustomConstructorIEnumerableInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomConstructorIEnumerableIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Equal([1, 2, 3], row1.Value);

        var row2 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Equal([1, -1, 2], row2.Value);

        var row3 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Equal([1], row3.Value);

        var row4 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<CustomConstructorIEnumerableIntClass>();
        Assert.Equal([-2], row5.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedCustomAddIEnumerableInt_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<CustomCustomAddIEnumerableIntClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Equal(new int[] { 1, 2, 3 }, row1.Value);

        var row2 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Equal(new int[] { 1, -1, 2 }, row2.Value);

        var row3 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Equal(new int[] { 1 }, row3.Value);

        var row4 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Empty(row4.Value);

        var row5 = sheet.ReadRow<CustomAddIEnumerableIntClass>();
        Assert.Equal(new int[] { -2 }, row5.Value);
    }

    public class ObjectArrayClass
    {
        public object?[] Value { get; set; } = default!;
    }

    public class DefaultObjectArrayClassMap : ExcelClassMap<ObjectArrayClass>
    {
        public DefaultObjectArrayClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class StringArrayClass
    {
        public string?[] Value { get; set; } = default!;
    }

    public class DefaultStringArrayClassMap : ExcelClassMap<StringArrayClass>
    {
        public DefaultStringArrayClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class IntArrayClass
    {
        public int[] Value { get; set; } = default!;
    }

    public class DefaultIntArrayClassMap : ExcelClassMap<IntArrayClass>
    {
        public DefaultIntArrayClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class CustomIntArrayClassMap : ExcelClassMap<IntArrayClass>
    {
        public CustomIntArrayClassMap()
        {
            Map(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class IEnumerableClass
    {
        public IEnumerable Value { get; set; } = default!;
    }

    public class DefaultIEnumerableClassMap : ExcelClassMap<IEnumerableClass>
    {
        public DefaultIEnumerableClassMap()
        {
            MapList<object>(p => p.Value);
        }
    }

    public class IEnumerableObjectClass
    {
        public IEnumerable<object?> Value { get; set; } = default!;
    }

    public class DefaultIEnumerableObjectClassMap : ExcelClassMap<IEnumerableObjectClass>
    {
        public DefaultIEnumerableObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class IEnumerableIntClass
    {
        public IEnumerable<int> Value { get; set; } = default!;
    }

    public class CustomConstructorIEnumerableIntClassMap : ExcelClassMap<IEnumerableIntClass>
    {
        public CustomConstructorIEnumerableIntClassMap()
        {
            Map(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class DefaultIEnumerableIntClassMap : ExcelClassMap<IEnumerableIntClass>
    {
        public DefaultIEnumerableIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class ICollectionClass
    {
        public ICollection Value { get; set; } = default!;
    }

    public class DefaultICollectionClassMap : ExcelClassMap<ICollectionClass>
    {
        public DefaultICollectionClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class ICollectionObjectClass
    {
        public ICollection<object?> Value { get; set; } = default!;
    }

    public class DefaultICollectionObjectClassMap : ExcelClassMap<ICollectionObjectClass>
    {
        public DefaultICollectionObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class ICollectionIntClass
    {
        public ICollection<int> Value { get; set; } = default!;
    }

    public class DefaultICollectionIntClassMap : ExcelClassMap<ICollectionIntClass>
    {
        public DefaultICollectionIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class CustomICollectionIntClassMap : ExcelClassMap<ICollectionIntClass>
    {
        public CustomICollectionIntClassMap()
        {
            Map(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class IListClass
    {
        public IList Value { get; set; } = default!;
    }

    public class DefaultIListClassMap : ExcelClassMap<IListClass>
    {
        public DefaultIListClassMap()
        {
            MapList<object>(p => p.Value);
        }
    }

    public class IListObjectClass
    {
        public IList<object?> Value { get; set; } = default!;
    }

    public class DefaultIListObjectClassMap : ExcelClassMap<IListObjectClass>
    {
        public DefaultIListObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class IListIntClass
    {
        public IList<int> Value { get; set; } = default!;
    }

    public class DefaultIListIntClassMap : ExcelClassMap<IListIntClass>
    {
        public DefaultIListIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class CustomIListIntClassMap : ExcelClassMap<IListIntClass>
    {
        public CustomIListIntClassMap()
        {
            Map(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class IReadOnlyCollectionObjectClass
    {
        public IReadOnlyCollection<object?> Value { get; set; } = default!;
    }

    public class DefaultIReadOnlyCollectionObjectClassMap : ExcelClassMap<IReadOnlyCollectionObjectClass>
    {
        public DefaultIReadOnlyCollectionObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class IReadOnlyCollectionIntClass
    {
        public IReadOnlyCollection<int> Value { get; set; } = default!;
    }

    public class DefaultIReadOnlyCollectionIntClassMap : ExcelClassMap<IReadOnlyCollectionIntClass>
    {
        public DefaultIReadOnlyCollectionIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class CustomIReadOnlyCollectionIntClassMap : ExcelClassMap<IReadOnlyCollectionIntClass>
    {
        public CustomIReadOnlyCollectionIntClassMap()
        {
            Map(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class IReadOnlyListObjectClass
    {
        public IReadOnlyList<object?> Value { get; set; } = default!;
    }

    public class DefaultIReadOnlyListObjectClassMap : ExcelClassMap<IReadOnlyListObjectClass>
    {
        public DefaultIReadOnlyListObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class IReadOnlyListIntClass
    {
        public IReadOnlyList<int> Value { get; set; } = default!;
    }

    public class DefaultIReadOnlyListIntClassMap : ExcelClassMap<IReadOnlyListIntClass>
    {
        public DefaultIReadOnlyListIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class CustomIReadOnlyListIntClassMap : ExcelClassMap<IReadOnlyListIntClass>
    {
        public CustomIReadOnlyListIntClassMap()
        {
            Map(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ArrayListClass
    {
        public ArrayList Value { get; set; } = default!;
    }

    public class DefaultArrayListClassMap : ExcelClassMap<ArrayListClass>
    {
        public DefaultArrayListClassMap()
        {
            MapList<string>(p => p.Value);
        }
    }

    public class ListObjectClass
    {
        public List<object?> Value { get; set; } = default!;
    }

    public class DefaultListObjectClassMap : ExcelClassMap<ListObjectClass>
    {
        public DefaultListObjectClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class ListIntClass
    {
        public List<int> Value { get; set; } = default!;
    }

    public class DefaultListIntClassMap : ExcelClassMap<ListIntClass>
    {
        public DefaultListIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class CustomListIntClassMap : ExcelClassMap<ListIntClass>
    {
        public CustomListIntClassMap()
        {
            Map(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ObservableCollectionIntClass
    {
        public ObservableCollection<int> Value { get; set; } = default!;
    }

    public class DefaultObservableCollectionIntClassMap : ExcelClassMap<ObservableCollectionIntClass>
    {
        public DefaultObservableCollectionIntClassMap()
        {
            Map(p => p.Value);
        }
    }

    public class CustomObservableCollectionIntClassMap : ExcelClassMap<ObservableCollectionIntClass>
    {
        public CustomObservableCollectionIntClassMap()
        {
            Map(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class QueueIntClass
    {
        public Queue<int> Value { get; set; } = default!;
    }

    public class DefaultQueueIntClassMap : ExcelClassMap<QueueIntClass>
    {
        public DefaultQueueIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomQueueIntClassMap : ExcelClassMap<QueueIntClass>
    {
        public CustomQueueIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class StackIntClass
    {
        public Stack<int> Value { get; set; } = default!;
    }

    public class DefaultStackIntClassMap : ExcelClassMap<StackIntClass>
    {
        public DefaultStackIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomStackIntClassMap : ExcelClassMap<StackIntClass>
    {
        public CustomStackIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class SortedSetIntClass
    {
        public SortedSet<int> Value { get; set; } = default!;
    }

    public class DefaultSortedSetIntClassMap : ExcelClassMap<SortedSetIntClass>
    {
        public DefaultSortedSetIntClassMap()
        {
            Map(p => (ICollection<int>)p.Value);
        }
    }

    public class CustomSortedSetIntClassMap : ExcelClassMap<SortedSetIntClass>
    {
        public CustomSortedSetIntClassMap()
        {
            Map(p => (ICollection<int>)p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class HashSetIntClass
    {
        public HashSet<int> Value { get; set; } = default!;
    }

    public class DefaultHashSetIntClassMap : ExcelClassMap<HashSetIntClass>
    {
        public DefaultHashSetIntClassMap()
        {
            Map(p => (ICollection<int>)p.Value);
        }
    }

    public class CustomHashSetIntClassMap : ExcelClassMap<HashSetIntClass>
    {
        public CustomHashSetIntClassMap()
        {
            Map(p => (ICollection<int>)p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class IImmutableListIntClass
    {
        public IImmutableList<int> Value { get; set; } = default!;
    }

    public class DefaultIImmutableListIntClassMap : ExcelClassMap<IImmutableListIntClass>
    {
        public DefaultIImmutableListIntClassMap()
        {
            Map(p => (IList<int>)p.Value);
        }
    }

    public class CustomIImmutableListIntClassMap : ExcelClassMap<IImmutableListIntClass>
    {
        public CustomIImmutableListIntClassMap()
        {
            Map(p => (IList<int>)p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class IImmutableStackIntClass
    {
        public IImmutableStack<int> Value { get; set; } = default!;
    }

    public class DefaultIImmutableStackIntClassMap : ExcelClassMap<IImmutableStackIntClass>
    {
        public DefaultIImmutableStackIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomIImmutableStackIntClassMap : ExcelClassMap<IImmutableStackIntClass>
    {
        public CustomIImmutableStackIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class IImmutableQueueIntClass
    {
        public IImmutableQueue<int> Value { get; set; } = default!;
    }

    public class DefaultIImmutableQueueIntClassMap : ExcelClassMap<IImmutableQueueIntClass>
    {
        public DefaultIImmutableQueueIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomIImmutableQueueIntClassMap : ExcelClassMap<IImmutableQueueIntClass>
    {
        public CustomIImmutableQueueIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class IImmutableSetIntClass
    {
        public IImmutableSet<int> Value { get; set; } = default!;
    }

    public class DefaultIImmutableSetIntClassMap : ExcelClassMap<IImmutableSetIntClass>
    {
        public DefaultIImmutableSetIntClassMap()
        {
            Map(p => (IList<int>)p.Value);
        }
    }

    public class CustomIImmutableSetIntClassMap : ExcelClassMap<IImmutableSetIntClass>
    {
        public CustomIImmutableSetIntClassMap()
        {
            Map(p => (IList<int>)p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ImmutableArrayIntClass
    {
        public ImmutableArray<int> Value { get; set; }
    }

    public class DefaultImmutableArrayIntClassMap : ExcelClassMap<ImmutableArrayIntClass>
    {
        public DefaultImmutableArrayIntClassMap()
        {
            Map(p => (IList<int>)p.Value);
        }
    }

    public class CustomImmutableArrayIntClassMap : ExcelClassMap<ImmutableArrayIntClass>
    {
        public CustomImmutableArrayIntClassMap()
        {
            Map(p => (IList<int>)p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ImmutableListIntClass
    {
        public ImmutableList<int> Value { get; set; } = default!;
    }

    public class DefaultImmutableListIntClassMap : ExcelClassMap<ImmutableListIntClass>
    {
        public DefaultImmutableListIntClassMap()
        {
            Map(p => (IList<int>)p.Value);
        }
    }

    public class CustomImmutableListIntClassMap : ExcelClassMap<ImmutableListIntClass>
    {
        public CustomImmutableListIntClassMap()
        {
            Map(p => (IList<int>)p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ImmutableStackIntClass
    {
        public ImmutableStack<int> Value { get; set; } = default!;
    }

    public class DefaultImmutableStackIntClassMap : ExcelClassMap<ImmutableStackIntClass>
    {
        public DefaultImmutableStackIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomImmutableStackIntClassMap : ExcelClassMap<ImmutableStackIntClass>
    {
        public CustomImmutableStackIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ImmutableQueueIntClass
    {
        public ImmutableQueue<int> Value { get; set; } = default!;
    }

    public class DefaultImmutableQueueIntClassMap : ExcelClassMap<ImmutableQueueIntClass>
    {
        public DefaultImmutableQueueIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomImmutableQueueIntClassMap : ExcelClassMap<ImmutableQueueIntClass>
    {
        public CustomImmutableQueueIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ImmutableSortedSetIntClass
    {
        public ImmutableSortedSet<int> Value { get; set; } = default!;
    }

    public class DefaultImmutableSortedSetIntClassMap : ExcelClassMap<ImmutableSortedSetIntClass>
    {
        public DefaultImmutableSortedSetIntClassMap()
        {
            Map(p => (IList<int>)p.Value);
        }
    }

    public class CustomImmutableSortedSetIntClassMap : ExcelClassMap<ImmutableSortedSetIntClass>
    {
        public CustomImmutableSortedSetIntClassMap()
        {
            Map(p => (IList<int>)p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ImmutableHashSetIntClass
    {
        public ImmutableHashSet<int> Value { get; set; } = default!;
    }

    public class DefaultImmutableHashSetIntClassMap : ExcelClassMap<ImmutableHashSetIntClass>
    {
        public DefaultImmutableHashSetIntClassMap()
        {
            Map(p => (ICollection<int>)p.Value);
        }
    }

    public class CustomImmutableHashSetIntClassMap : ExcelClassMap<ImmutableHashSetIntClass>
    {
        public CustomImmutableHashSetIntClassMap()
        {
            Map(p => (ICollection<int>)p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ConcurrentQueueIntClass
    {
        public ConcurrentQueue<int> Value { get; set; } = default!;
    }

    public class DefaultConcurrentQueueIntClassMap : ExcelClassMap<ConcurrentQueueIntClass>
    {
        public DefaultConcurrentQueueIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomConcurrentQueueIntClassMap : ExcelClassMap<ConcurrentQueueIntClass>
    {
        public CustomConcurrentQueueIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ConcurrentStackIntClass
    {
        public ConcurrentStack<int> Value { get; set; } = default!;
    }

    public class DefaultConcurrentStackIntClassMap : ExcelClassMap<ConcurrentStackIntClass>
    {
        public DefaultConcurrentStackIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomConcurrentStackIntClassMap : ExcelClassMap<ConcurrentStackIntClass>
    {
        public CustomConcurrentStackIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class ConcurrentBagIntClass
    {
        public ConcurrentBag<int> Value { get; set; } = default!;
    }

    public class DefaultConcurrentBagIntClassMap : ExcelClassMap<ConcurrentBagIntClass>
    {
        public DefaultConcurrentBagIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomConcurrentBagIntClassMap : ExcelClassMap<ConcurrentBagIntClass>
    {
        public CustomConcurrentBagIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class BlockingCollectionIntClass
    {
        public BlockingCollection<int> Value { get; set; } = default!;
    }

    public class DefaultBlockingCollectionIntClassMap : ExcelClassMap<BlockingCollectionIntClass>
    {
        public DefaultBlockingCollectionIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomBlockingCollectionIntClassMap : ExcelClassMap<BlockingCollectionIntClass>
    {
        public CustomBlockingCollectionIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class CustomConstructorIEnumerable<T> : IEnumerable<T>
    {
        private readonly IEnumerable<T> _inner;

        public CustomConstructorIEnumerable(IEnumerable<T> inner)
        {
            _inner = inner;
        }

        public IEnumerator<T> GetEnumerator() => _inner.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => ((IEnumerable)_inner).GetEnumerator();
    }

    public class CustomConstructorIEnumerableIntClass
    {
        public CustomConstructorIEnumerable<int> Value { get; set; } = default!;
    }

    public class DefaultCustomConstructorIEnumerableIntClassMap : ExcelClassMap<CustomConstructorIEnumerableIntClass>
    {
        public DefaultCustomConstructorIEnumerableIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomCustomConstructorIEnumerableIntClassMap : ExcelClassMap<CustomConstructorIEnumerableIntClass>
    {
        public CustomCustomConstructorIEnumerableIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    public class CustomAddIEnumerable<T> : IEnumerable<T>
    {
        private readonly List<T> _inner = new();

        public IEnumerator<T> GetEnumerator() => _inner.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => ((IEnumerable)_inner).GetEnumerator();

        public void Add(T value) => _inner.Add(value);
    }

    public class CustomAddIEnumerableIntClass
    {
        public CustomAddIEnumerable<int> Value { get; set; } = default!;
    }

    public class DefaultCustomAddIEnumerableIntClassMap : ExcelClassMap<CustomAddIEnumerableIntClass>
    {
        public DefaultCustomAddIEnumerableIntClassMap()
        {
            Map<int>(p => p.Value);
        }
    }

    public class CustomCustomAddIEnumerableIntClassMap : ExcelClassMap<CustomAddIEnumerableIntClass>
    {
        public CustomCustomAddIEnumerableIntClassMap()
        {
            Map<int>(p => p.Value)
                .WithElementMap(p => p
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );
        }
    }

    [Fact]
    public void ReadRow_MultiMapMissingRow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<DefaultMissingColumnClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnClass>());
    }

    [Fact]
    public void ReadRow_MultiMapOptionalMissingRow_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<OptionalMissingColumnClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row = sheet.ReadRow<MissingColumnClass>();
        Assert.Null(row.MissingColumn);
    }

    public class MissingColumnClass
    {
        public int[] MissingColumn { get; set; } = default!;
    }

    private class DefaultMissingColumnClassMap : ExcelClassMap<MissingColumnClass>
    {
        public DefaultMissingColumnClassMap()
        {
            Map(p => p.MissingColumn);
        }
    }

    private class OptionalMissingColumnClassMap : ExcelClassMap<MissingColumnClass>
    {
        public OptionalMissingColumnClassMap()
        {
            Map(p => p.MissingColumn)
                .MakeOptional();
        }
    }

    [Fact]
    public void ReadRow_AutoMappedMultiImmutableArrayBuilder_ThrowExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ImmutableArrayBuilderIntClass>());
    }

    public class ImmutableArrayBuilderIntClass
    {
        public ImmutableArray<int>.Builder Value { get; set; } = default!;
    }
}
