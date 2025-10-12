using System;
using System.Collections;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Xunit;

namespace ExcelMapper.Factories;

public class AddEnumerableFactoryTests
{
    [Theory]
    [InlineData(typeof(List<int>))]
    [InlineData(typeof(ICollectionGeneric<int>))]
    [InlineData(typeof(AddClass<int>))]
    [InlineData(typeof(ArrayList))]
    [InlineData(typeof(HashSet<int>))]
    [InlineData(typeof(Collection<int>))]
    [InlineData(typeof(ObservableCollection<int>))]
    public void Ctor_Type(Type collectionType)
    {
        var factory = new AddEnumerableFactory<int>(collectionType);
        Assert.Equal(collectionType, factory.CollectionType);
    }

    [Fact]
    public void Ctor_NullCollectionType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("collectionType", () => new AddEnumerableFactory<int>(null!));
    }

    [Theory]
    [InlineData(typeof(IEnumerable))]
    [InlineData(typeof(IEnumerable<int>))]
    [InlineData(typeof(IEnumerableGeneric<int>))]
    [InlineData(typeof(ICollection))]
    [InlineData(typeof(ICollection<int>))]
    [InlineData(typeof(ICollectionNonGeneric))]
    [InlineData(typeof(int[]))]
    [InlineData(typeof(List<string>))]
    [InlineData(typeof(AddClass<string>))]
    [InlineData(typeof(Collection<string>))]
    [InlineData(typeof(ObservableCollection<string>))]
    [InlineData(typeof(ReadOnlyCollection<int>))]
    [InlineData(typeof(ReadOnlyObservableCollection<int>))]
    [InlineData(typeof(ReadOnlyObservableCollection<string>))]
    [InlineData(typeof(AbstractClass))]
    [InlineData(typeof(FrozenSet<int>))]
    [InlineData(typeof(ReadOnlySet<int>))]
    public void Ctor_InvalidCollectionType_ThrowsArgumentException(Type collectionType)
    {
        Assert.Throws<ArgumentException>("collectionType", () => new AddEnumerableFactory<int>(collectionType));
    }

    [Fact]
    public void Begin_End_Success()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));

        // Begin.
        factory.Begin(1);
        factory.Add(1);
        var value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([1], value);

        // Begin again.
        factory.Begin(1);
        factory.Add(2);
        value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([2], value);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        factory.Add(2);

        factory.Add(3);

        var value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([2, 3], value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        Assert.Throws<ExcelMappingException>(() => factory.Add(1));
    }

    [Fact]
    public void Add_Throws_Rethrows()
    {
        var factory = new AddEnumerableFactory<int>(typeof(AddThrowsClass<int>));
        factory.Begin(1);
        Assert.Throws<InvalidOperationException>(() => factory.Add(1));
    }

    private class AddThrowsClass<T>
    {
        public void Add(T item) => throw new InvalidOperationException();
    }

    [Fact]
    public void Set_Invoke_ThrowsNotSupportedException()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        Assert.Throws<NotSupportedException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        Assert.Throws<ExcelMappingException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new AddEnumerableFactory<int>(typeof(List<int>));
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([], value);
    }

    private abstract class AbstractClass
    {
        public void Add(int value) => throw new NotImplementedException();
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

    private class AddClass<T>
    {
        public void Add(T item) => throw new NotImplementedException();
    }
}
