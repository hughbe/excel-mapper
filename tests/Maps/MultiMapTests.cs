using System.Collections;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MultiMapTests
    {
        [Fact]
        public void ReadRow_MultiMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("MultiMap.xlsx"))
            {
                importer.Configuration.RegisterClassMap(new MultiMapRowMap());

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                MultiMapRow row1 = sheet.ReadRow<MultiMapRow>();
                Assert.Equal(new int[] { 1, 2, 3 }, row1.MultiMapName);
                Assert.Equal(new string[] { "a", "b" }, row1.MultiMapIndex);
                Assert.Equal(new int[] { 1, 2 }, row1.IEnumerableInt);
                Assert.Equal(new bool[] { true, false }, row1.ICollectionBool);
                Assert.Equal(new string[] { "a", "b" }, row1.IListString);
                Assert.Equal(new string[] { "1", "2" }, row1.ListString);
                Assert.Equal(new string[] { "1", "2" }, row1._concreteICollection);

                MultiMapRow row2 = sheet.ReadRow<MultiMapRow>();
                Assert.Equal(new int[] { 1, -1, 3 }, row2.MultiMapName);
                Assert.Equal(new string[] { null, null }, row2.MultiMapIndex);
                Assert.Equal(new int[] { 0, 0 }, row2.IEnumerableInt);
                Assert.Equal(new bool[] { false, true }, row2.ICollectionBool);
                Assert.Equal(new string[] { "c", "d" }, row2.IListString);
                Assert.Equal(new string[] { "3", "4" }, row2.ListString);
                Assert.Equal(new string[] { "3", "4" }, row2._concreteICollection);

                MultiMapRow row3 = sheet.ReadRow<MultiMapRow>();
                Assert.Equal(new int[] { -1, -1, -1 }, row3.MultiMapName);
                Assert.Equal(new string[] { null, "d" }, row3.MultiMapIndex);
                Assert.Equal(new int[] { 5, 6 }, row3.IEnumerableInt);
                Assert.Equal(new bool[] { false, false }, row3.ICollectionBool);
                Assert.Equal(new string[] { "e", "f" }, row3.IListString);
                Assert.Equal(new string[] { "5", "6" }, row3.ListString);
                Assert.Equal(new string[] { "5", "6" }, row3._concreteICollection);

                MultiMapRow row4 = sheet.ReadRow<MultiMapRow>();
                Assert.Equal(new int[] { -2, -2, 3 }, row4.MultiMapName);
                Assert.Equal(new string[] { "d", null }, row4.MultiMapIndex);
                Assert.Equal(new int[] { 7, 8 }, row4.IEnumerableInt);
                Assert.Equal(new bool[] { false, true }, row4.ICollectionBool);
                Assert.Equal(new string[] { "g", "h" }, row4.IListString);
                Assert.Equal(new string[] { "7", "8" }, row4.ListString);
                Assert.Equal(new string[] { "7", "8" }, row4._concreteICollection);
            }
        }

        private class MultiMapRow
        {
            public int[] MultiMapName { get; set; }
            public CustomList MultiMapIndex { get; set; }
            public IEnumerable<int> IEnumerableInt { get; set; }
            public ICollection<bool> ICollectionBool { get; set; }
            public IList<string> IListString { get; set; }
            public List<string> ListString { get; set; }
#pragma warning disable 0649
            public SortedSet<string> _concreteICollection;
#pragma warning restore 0649
        }

        private class MultiMapRowMap : ExcelClassMap<MultiMapRow>
        {
            public MultiMapRowMap()
            {
                Map(p => p.MultiMapName)
                    .WithColumnNames("MultiMapName1", "MultiMapName2", "MultiMapName3")
                    .WithElementMap(e => e
                        .WithEmptyFallback(-1)
                        .WithInvalidFallback(-2)
                    );

                Map<string>(p => p.MultiMapIndex)
                    .WithColumnIndices(3, 4);

                Map(p => p.IEnumerableInt)
                    .WithColumnNames(new List<string> { "IEnumerableInt1", "IEnumerableInt2" })
                    .WithElementMap(e => e
                        .WithValueFallback(default(int))
                    );

                Map(p => p.ICollectionBool)
                    .WithColumnIndices(new List<int> { 7, 8 })
                    .WithElementMap(e => e
                        .WithValueFallback(default(bool))
                    );

                Map(p => p.IListString)
                    .WithColumnNames("IListString1", "IListString2");

                Map(p => p.ListString)
                    .WithColumnNames("ListString1", "ListString2");

                Map<string>(p => p._concreteICollection)
                    .WithColumnNames("ListString1", "ListString2");
            }
        }

        public interface INonGenericInteface { }
        public interface IGenericInterface<T> { }
        public interface IMultipleGenericInterface<T, U>{ }

        public class CustomList : INonGenericInteface, IGenericInterface<CustomList>, IList<string>, IMultipleGenericInterface<string, int>
        {
            private IList<string> Inner { get; } = new List<string>();

            public string this[int index]
            {
                get => Inner[0];
                set => Inner[0] = value;
            }

            public int Count => Inner.Count;

            public bool IsReadOnly => Inner.IsReadOnly;

            public void Add(string item) => Inner.Add(item);

            public void Clear() => Inner.Clear();

            public bool Contains(string item) => Inner.Contains(item);

            public void CopyTo(string[] array, int arrayIndex) => Inner.CopyTo(array, arrayIndex);

            public IEnumerator<string> GetEnumerator() => Inner.GetEnumerator();

            public int IndexOf(string item) => Inner.IndexOf(item);

            public void Insert(int index, string item) => Inner.Insert(index, item);

            public bool Remove(string item) => Inner.Remove(item);

            public void RemoveAt(int index) => Inner.RemoveAt(index);

            IEnumerator IEnumerable.GetEnumerator() => Inner.GetEnumerator();
        }
    }
}
