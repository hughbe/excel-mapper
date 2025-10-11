using System;
using System.Collections;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Factories;

public class ConstructorEnumerableFactoryTests
{
    [Theory]
    [InlineData(typeof(List<int>))]
    [InlineData(typeof(ConstructorIEnumerable))]
    [InlineData(typeof(ConstructorIEnumerableT<int>))]
    [InlineData(typeof(ConstructorICollection))]
    [InlineData(typeof(ArrayList))]
    public void Ctor_Type(Type collectionType)
    {
        var factory = new ConstructorEnumerableFactory<int>(collectionType);
        Assert.Equal(collectionType, factory.CollectionType);
    }

    [Fact]
    public void Ctor_NullCollectionType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("collectionType", () => new ConstructorEnumerableFactory<int>(null!));
    }

    [Theory]
    [InlineData(typeof(IEnumerable))]
    [InlineData(typeof(IEnumerable<int>))]
    [InlineData(typeof(ICollection))]
    [InlineData(typeof(ICollection<int>))]
    [InlineData(typeof(int[]))]
    [InlineData(typeof(NoConstructor))]
    [InlineData(typeof(List<string>))]
    [InlineData(typeof(ConstructorIEnumerableT<string>))]
    [InlineData(typeof(ConstructorICollectionT<int>))]
    [InlineData(typeof(ConstructorICollectionT<string>))]
    [InlineData(typeof(AbstractClass))]
    public void Ctor_InvalidCollectionType_ThrowsArgumentException(Type collectionType)
    {
        Assert.Throws<ArgumentException>("collectionType", () => new ConstructorEnumerableFactory<int>(collectionType));
    }

    [Fact]
    public void Begin_End_Success()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));

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
    public void Begin_EndIEnumerableInt_Success()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(ConstructorIEnumerableT<int>));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<ConstructorIEnumerableT<int>>(factory.End());
        Assert.Equal([], value.Value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<ConstructorIEnumerableT<int>>(factory.End());
        Assert.Equal([], value.Value);
    }

    [Fact]
    public void Begin_EndICollection_Success()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(ConstructorICollection));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<ConstructorICollection>(factory.End());
        Assert.Empty(value.Value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<ConstructorICollection>(factory.End());
        Assert.Empty(value.Value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));

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
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        factory.Add(2);

        factory.Add(3);

        var value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([2, 3], value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        Assert.Throws<ExcelMappingException>(() => factory.Add(1));
    }

    [Fact]
    public void Set_Invoke_Success()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);

        factory.Set(0, 1);
        Assert.Equal([1], Assert.IsType<List<int>>(factory.End()));
    }

    [Fact]
    public void Set_InvokeOutOfRange_Success()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);

        factory.Set(0, 1);
        factory.Set(5, 2);
        Assert.Equal([1, 0, 0, 0, 0, 2], Assert.IsType<List<int>>(factory.End()));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        Assert.Throws<ExcelMappingException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void Set_NegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        Assert.Throws<ArgumentOutOfRangeException>("index", () => factory.Set(-1, 1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
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
        var factory = new ConstructorEnumerableFactory<int>(typeof(List<int>));
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<List<int>>(factory.End());
        Assert.Equal([], value);
    }

    private abstract class AbstractClass
    {
        public IEnumerable<int> Values { get; } = default!;

        public AbstractClass(IEnumerable<int> values)
        {
            Values = values;
        }
    }

    private class NoConstructor
    {
    }

    private class ConstructorIEnumerable
    {
        public IEnumerable Value { get; }

        public ConstructorIEnumerable(IEnumerable value) => Value = value;
    }

    private class ConstructorIEnumerableT<T>
    {
        public IEnumerable<T> Value { get; }

        public ConstructorIEnumerableT(IEnumerable<T> value) => Value = value;
    }

    private class ConstructorICollection
    {
        public ICollection Value { get; }

        public ConstructorICollection(ICollection value) => Value = value;
    }

    private class ConstructorICollectionT<T>
    {
        public ICollection<T> Value { get; }

        public ConstructorICollectionT(ICollection<T> value) => Value = value;
    }
}
