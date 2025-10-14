using System;
using System.Collections;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Xunit;

namespace ExcelMapper.Factories;

public class ConstructorSetEnumerableFactoryTests
{
    [Theory]
    [InlineData(typeof(List<int>))]
    [InlineData(typeof(ConstructorIEnumerable))]
    [InlineData(typeof(ConstructorIEnumerableT<int>))]
    [InlineData(typeof(ConstructorICollectionT<int>))]
    [InlineData(typeof(ConstructorISetT<int>))]
    [InlineData(typeof(ObservableCollection<int>))]
    [InlineData(typeof(HashSet<int>))]
    [InlineData(typeof(ReadOnlySet<int>))]
    public void Ctor_Type(Type setType)
    {
        var factory = new ConstructorSetEnumerableFactory<int>(setType);
        Assert.Equal(setType, factory.SetType);
    }

    [Fact]
    public void Ctor_NullSetType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("setType", () => new ConstructorSetEnumerableFactory<int>(null!));
    }

    [Theory]
    [InlineData(typeof(IEnumerable))]
    [InlineData(typeof(IEnumerable<int>))]
    [InlineData(typeof(ICollection))]
    [InlineData(typeof(ICollection<int>))]
    [InlineData(typeof(IList))]
    [InlineData(typeof(IList<int>))]
    [InlineData(typeof(int[]))]
    [InlineData(typeof(KeyValuePair<string, int>[]))]
    [InlineData(typeof(NoConstructor))]
    [InlineData(typeof(List<string>))]
    [InlineData(typeof(ConstructorIList))]
    [InlineData(typeof(ConstructorIEnumerableT<string>))]
    [InlineData(typeof(ConstructorICollection))]
    [InlineData(typeof(ConstructorICollectionT<string>))]
    [InlineData(typeof(ConstructorIListT<int>))]
    [InlineData(typeof(ConstructorIListT<string>))]
    [InlineData(typeof(ConstructorISetT<string>))]
    [InlineData(typeof(ConstructorIReadOnlySetT<int>))]
    [InlineData(typeof(ConstructorIReadOnlySetT<string>))]
    [InlineData(typeof(ArrayList))]
    [InlineData(typeof(Queue))]
    [InlineData(typeof(Stack))]
    [InlineData(typeof(AbstractClass))]
    [InlineData(typeof(Collection<int>))]
    [InlineData(typeof(ReadOnlyObservableCollection<int>))]
    [InlineData(typeof(ReadOnlyObservableCollection<string>))]
    [InlineData(typeof(ObservableCollection<string>))]
    [InlineData(typeof(ReadOnlyCollection<int>))]
    [InlineData(typeof(ReadOnlyCollection<string>))]
    [InlineData(typeof(FrozenSet<string>))]
    [InlineData(typeof(HashSet<string>))]
    [InlineData(typeof(ReadOnlySet<string>))]
    public void Ctor_InvalidSetType_ThrowsArgumentException(Type setType)
    {
        Assert.Throws<ArgumentException>("setType", () => new ConstructorSetEnumerableFactory<int>(setType));
    }

    [Fact]
    public void Begin_End_Success()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<ReadOnlySet<int>>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<ReadOnlySet<int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Begin_NegativeCount_ThrowsArgumentOutOfRangeException()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        Assert.Throws<ArgumentOutOfRangeException>("count", () => factory.Begin(-1));
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));

        // Begin.
        factory.Begin(1);
        factory.Add(1);
        var value = Assert.IsType<ReadOnlySet<int>>(factory.End());
        Assert.Equal([1], value);

        // Begin again.
        factory.Begin(1);
        factory.Add(2);
        value = Assert.IsType<ReadOnlySet<int>>(factory.End());
        Assert.Equal([2], value);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        factory.Begin(1);
        factory.Add(2);

        factory.Add(3);
        
        var value = Assert.IsType<ReadOnlySet<int>>(factory.End());
        Assert.Equal([2, 3], value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        Assert.Throws<ExcelMappingException>(() => factory.Add(1));
    }

    [Fact]
    public void Set_Invoke_ThrowsNotSupportedException()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        factory.Begin(1);
        Assert.Throws<NotSupportedException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        Assert.Throws<ExcelMappingException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<ReadOnlySet<int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new ConstructorSetEnumerableFactory<int>(typeof(ReadOnlySet<int>));
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<ReadOnlySet<int>>(factory.End());
        Assert.Equal([], value);
    }


    private abstract class AbstractClass
    {
        public ISet<int> Values { get; } = default!;

        public AbstractClass(ISet<int> values)
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

    private class ConstructorIList
    {
        public IList Value { get; }

        public ConstructorIList(IList value) => Value = value;
    }

    private class ConstructorIListT<T>
    {
        public IList<T> Value { get; }

        public ConstructorIListT(IList<T> value) => Value = value;
    }

    private class ConstructorISetT<T>
    {
        public ISet<T> Value { get; }

        public ConstructorISetT(ISet<T> value) => Value = value;
    }

    private class ConstructorIReadOnlySetT<T>
    {
        public IReadOnlySet<T> Value { get; }

        public ConstructorIReadOnlySetT(IReadOnlySet<T> value) => Value = value;
    }
}
