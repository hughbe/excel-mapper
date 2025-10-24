using System.Collections;
using System.Collections.Frozen;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelMapper.Factories;

public class IListImplementingEnumerableFactoryTests
{
    [Theory]
    [InlineData(typeof(List<int>))]
    [InlineData(typeof(List<string>))]
    [InlineData(typeof(IListNonGeneric))]
    [InlineData(typeof(Collection<int>))]
    [InlineData(typeof(Collection<string>))]
    [InlineData(typeof(ArrayList))]
    [InlineData(typeof(SubCollectionBase))]
    public void Ctor_Type(Type listType)
    {
        var factory = new IListImplementingEnumerableFactory<int>(listType);
        Assert.Equal(listType, factory.ListType);
    }

    [Fact]
    public void Ctor_NullListType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("listType", () => new IListImplementingEnumerableFactory<int>(null!));
    }

    [Theory]
    [InlineData(typeof(IEnumerable))]
    [InlineData(typeof(IEnumerable<int>))]
    [InlineData(typeof(ICollection))]
    [InlineData(typeof(ICollection<int>))]
    [InlineData(typeof(IReadOnlyCollection<int>))]
    [InlineData(typeof(IList))]
    [InlineData(typeof(IList<int>))]
    [InlineData(typeof(IReadOnlyList<int>))]
    [InlineData(typeof(int[]))]
    [InlineData(typeof(Stack))]
    [InlineData(typeof(Queue))]
    [InlineData(typeof(IListGeneric<int>))]
    [InlineData(typeof(ICollectionNonGeneric))]
    [InlineData(typeof(ICollectionGeneric<int>))]
    [InlineData(typeof(AbstractClass))]
    [InlineData(typeof(ImmutableList<int>))]
    [InlineData(typeof(FrozenSet<int>))]
#if NET9_0_OR_GREATER
    [InlineData(typeof(ReadOnlySet<int>))]
#endif
    [InlineData(typeof(HashSet<int>))]
    [InlineData(typeof(CollectionBase))]
    [InlineData(typeof(ReadOnlyCollection<int>))]
    [InlineData(typeof(ReadOnlyObservableCollection<int>))]
    [InlineData(typeof(NoConstructorClass))]
    [InlineData(typeof(NoConstructorClass<int>))]
    public void Ctor_InvalidListType_ThrowsArgumentException(Type listType)
    {
        Assert.Throws<ArgumentException>("listType", () => new IListImplementingEnumerableFactory<int>(listType));
    }

    [Fact]
    public void Begin_End_Success()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<SubCollectionBase>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<SubCollectionBase>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Begin_NegativeCount_ThrowsArgumentOutOfRangeException()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        Assert.Throws<ArgumentOutOfRangeException>("count", () => factory.Begin(-1));
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));

        // Begin.
        factory.Begin(1);
        factory.Add(1);
        var value = Assert.IsType<SubCollectionBase>(factory.End());
        Assert.Equal([1], value.Cast<int>());

        // Begin again.
        factory.Begin(1);
        factory.Add(2);
        value = Assert.IsType<SubCollectionBase>(factory.End());
        Assert.Equal([2], value.Cast<int>());
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        factory.Begin(1);
        factory.Add(2);

        factory.Add(3);

        var value = Assert.IsType<SubCollectionBase>(factory.End());
        Assert.Equal([2, 3], value.Cast<int>());
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        Assert.Throws<ExcelMappingException>(() => factory.Add(1));
    }

    [Fact]
    public void Set_Invoke_Success()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        factory.Begin(1);

        factory.Set(0, 1);
        Assert.Equal([1], Assert.IsType<SubCollectionBase>(factory.End()).Cast<int>());
    }

    [Fact]
    public void Set_InvokeOutOfRange_Success()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        factory.Begin(1);

        factory.Set(0, 1);
        factory.Set(5, 2);
        Assert.Equal([1, 0, 0, 0, 0, 2], Assert.IsType<SubCollectionBase>(factory.End()).Cast<int>());
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        Assert.Throws<ExcelMappingException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void Set_NegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        Assert.Throws<ArgumentOutOfRangeException>("index", () => factory.Set(-1, 1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<SubCollectionBase>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new IListImplementingEnumerableFactory<int>(typeof(SubCollectionBase));
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<SubCollectionBase>(factory.End());
        Assert.Equal([], value);
    }

    private abstract class AbstractClass
    {
    }

    private class IEnumerableNonGeneric : IEnumerable
    {
        public IEnumerator GetEnumerator() => throw new NotImplementedException();
    }

    private class IEnumerableGeneric<T> : IEnumerable<T>
    {
        public IEnumerator GetEnumerator() => throw new NotImplementedException();

        IEnumerator<T> IEnumerable<T>.GetEnumerator() => throw new NotImplementedException();
    }

    private class ICollectionNonGeneric : ICollection
    {
        public int Count => throw new NotImplementedException();

        public bool IsSynchronized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public void CopyTo(Array array, int index) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new NotImplementedException();
    }

    private class ICollectionGeneric<T> : ICollection<T>
    {
        public int Count => throw new NotImplementedException();

        public bool IsSynchronized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(T item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(T item) => throw new NotImplementedException();

        public void CopyTo(Array array, int index) => throw new NotImplementedException();

        public void CopyTo(T[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new NotImplementedException();

        public bool Remove(T item) => throw new NotImplementedException();

        IEnumerator<T> IEnumerable<T>.GetEnumerator() => throw new NotImplementedException();
    }

    private class IListNonGeneric : IList
    {
        public object? this[int index] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public int Count => throw new NotImplementedException();

        public bool IsSynchronized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public bool IsFixedSize => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public int Add(object? value) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(object? value) => throw new NotImplementedException();

        public void CopyTo(Array array, int index) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new NotImplementedException();

        public int IndexOf(object? value) => throw new NotImplementedException();

        public void Insert(int index, object? value) => throw new NotImplementedException();

        public void Remove(object? value) => throw new NotImplementedException();

        public void RemoveAt(int index) => throw new NotImplementedException();
    }

    private class IListGeneric<T> : IList<T>
    {
        public T this[int index] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public int Count => throw new NotImplementedException();

        public bool IsSynchronized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(T item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(T item) => throw new NotImplementedException();

        public void CopyTo(Array array, int index) => throw new NotImplementedException();

        public void CopyTo(T[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new NotImplementedException();

        public int IndexOf(T item) => throw new NotImplementedException();

        public void Insert(int index, T item) => throw new NotImplementedException();

        public bool Remove(T item) => throw new NotImplementedException();

        public void RemoveAt(int index) => throw new NotImplementedException();

        IEnumerator<T> IEnumerable<T>.GetEnumerator() => throw new NotImplementedException();
    }

    private class SubCollectionBase : CollectionBase
    {
        protected override void OnValidate(object value) { }
    }

    private class NoConstructorClass : IList
    {
        private NoConstructorClass()
        {
        }

        public object? this[int index] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public bool IsFixedSize => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public int Count => throw new NotImplementedException();

        public bool IsSynchronized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public int Add(object? value) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(object? value) => throw new NotImplementedException();

        public void CopyTo(Array array, int index) => throw new NotImplementedException();

        public IEnumerator GetEnumerator() => throw new NotImplementedException();

        public int IndexOf(object? value) => throw new NotImplementedException();

        public void Insert(int index, object? value) => throw new NotImplementedException();

        public void Remove(object? value) => throw new NotImplementedException();

        public void RemoveAt(int index) => throw new NotImplementedException();
    }

    private class NoConstructorClass<T> : IList<T>
    {
        private NoConstructorClass()
        {
        }

        public T this[int index] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(T item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(T item) => throw new NotImplementedException();

        public void CopyTo(T[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator<T> GetEnumerator() => throw new NotImplementedException();

        public int IndexOf(T item) => throw new NotImplementedException();

        public void Insert(int index, T item) => throw new NotImplementedException();

        public bool Remove(T item) => throw new NotImplementedException();

        public void RemoveAt(int index) => throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
