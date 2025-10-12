using System;
using System.Collections;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Xunit;

namespace ExcelMapper.Factories;

public class ISetTImplementingEnumerableFactoryTests
{
    [Theory]
    [InlineData(typeof(HashSet<int>))]
    [InlineData(typeof(ReadOnlySet<int>))]
    [InlineData(typeof(ISetGeneric<int>))]
    public void Ctor_Type(Type setType)
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(setType);
        Assert.Equal(setType, factory.SetType);
    }

    [Fact]
    public void Ctor_NullSetType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("setType", () => new ISetTImplementingEnumerableFactory<int>(null!));
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
    [InlineData(typeof(ArrayList))]
    [InlineData(typeof(IListNonGeneric))]
    [InlineData(typeof(ICollectionNonGeneric))]
    [InlineData(typeof(ICollectionGeneric<int>))]
    [InlineData(typeof(IListGeneric<int>))]
    [InlineData(typeof(List<int>))]
    [InlineData(typeof(List<string>))]
    [InlineData(typeof(AbstractClass))]
    [InlineData(typeof(FrozenSet<int>))]
    [InlineData(typeof(Collection<string>))]
    [InlineData(typeof(Collection<int>))]
    [InlineData(typeof(ReadOnlyCollection<string>))]
    [InlineData(typeof(ReadOnlyCollection<int>))]
    [InlineData(typeof(HashSet<string>))]
    [InlineData(typeof(ReadOnlySet<string>))]
    [InlineData(typeof(ISet<int>))]
    [InlineData(typeof(ISet<string>))]
    [InlineData(typeof(IReadOnlySet<int>))]
    [InlineData(typeof(IReadOnlySet<string>))]
    [InlineData(typeof(ISetGeneric<string>))]
    [InlineData(typeof(IReadOnlySetGeneric<int>))]
    [InlineData(typeof(IReadOnlySetGeneric<string>))]
    public void Ctor_InvalidSetType_ThrowsArgumentException(Type setType)
    {
        Assert.Throws<ArgumentException>("setType", () => new ISetTImplementingEnumerableFactory<int>(setType));
    }

    [Fact]
    public void Begin_End_Success()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<HashSet<int>>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<HashSet<int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }
    [Fact]
    public void Add_End_Success()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));

        // Begin.
        factory.Begin(1);
        factory.Add(1);
        var value = Assert.IsType<HashSet<int>>(factory.End());
        Assert.Equal([1], value);

        // Begin again.
        factory.Begin(1);
        factory.Add(2);
        value = Assert.IsType<HashSet<int>>(factory.End());
        Assert.Equal([2], value);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        factory.Begin(1);
        factory.Add(2);

        factory.Add(3);
        
        var value = Assert.IsType<HashSet<int>>(factory.End());
        Assert.Equal([2, 3], value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        Assert.Throws<ExcelMappingException>(() => factory.Add(1));
    }

    [Fact]
    public void Set_Invoke_ThrowsNotSupportedException()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        factory.Begin(1);
        Assert.Throws<NotSupportedException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        Assert.Throws<ExcelMappingException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<HashSet<int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new ISetTImplementingEnumerableFactory<int>(typeof(HashSet<int>));
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<HashSet<int>>(factory.End());
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

    private class IReadOnlySetGeneric<T> : IReadOnlySet<T>
    {
        public int Count => throw new NotImplementedException();

        public bool Contains(T item) =>throw new NotImplementedException();

        public IEnumerator<T> GetEnumerator() =>throw new NotImplementedException();

        public bool IsProperSubsetOf(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool IsProperSupersetOf(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool IsSubsetOf(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool IsSupersetOf(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool Overlaps(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool SetEquals(IEnumerable<T> other) =>throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private class ISetGeneric<T> : ISet<T>
    {
        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public bool Add(T item) =>throw new NotImplementedException();

        public void Clear() =>throw new NotImplementedException();

        public bool Contains(T item) =>throw new NotImplementedException();

        public void CopyTo(T[] array, int arrayIndex) =>throw new NotImplementedException();

        public void ExceptWith(IEnumerable<T> other) =>throw new NotImplementedException();

        public IEnumerator<T> GetEnumerator() =>throw new NotImplementedException();

        public void IntersectWith(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool IsProperSubsetOf(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool IsProperSupersetOf(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool IsSubsetOf(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool IsSupersetOf(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool Overlaps(IEnumerable<T> other) =>throw new NotImplementedException();

        public bool Remove(T item) =>throw new NotImplementedException();

        public bool SetEquals(IEnumerable<T> other) =>throw new NotImplementedException();

        public void SymmetricExceptWith(IEnumerable<T> other) =>throw new NotImplementedException();

        public void UnionWith(IEnumerable<T> other) =>throw new NotImplementedException();

        void ICollection<T>.Add(T item) =>throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
