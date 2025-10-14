using System;
using System.Collections;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using Xunit;

namespace ExcelMapper.Factories;

public class ConstructorDictionaryFactoryTests
{
    [Theory]
    [InlineData(typeof(ConstructorIEnumerable))]
    [InlineData(typeof(ConstructorIEnumerableT<KeyValuePair<string, int>>))]
    [InlineData(typeof(ConstructorICollectionT<KeyValuePair<string, int>>))]
    [InlineData(typeof(Dictionary<string, int>))]
    [InlineData(typeof(ReadOnlyDictionary<string, int>))]
    [InlineData(typeof(ConstructorIDictionaryT<string, int>))]
    public void Ctor_Type(Type dictionaryType)
    {
        var factory = new ConstructorDictionaryFactory<string, int>(dictionaryType);
        Assert.Equal(dictionaryType, factory.DictionaryType);
    }

    [Fact]
    public void Ctor_NullDictionaryType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("dictionaryType", () => new ConstructorDictionaryFactory<string, int>(null!));
    }

    [Theory]
    [InlineData(typeof(IEnumerable))]
    [InlineData(typeof(IEnumerable<int>))]
    [InlineData(typeof(ICollection))]
    [InlineData(typeof(ICollection<int>))]
    [InlineData(typeof(IList))]
    [InlineData(typeof(IList<int>))]
    [InlineData(typeof(int[]))]
    [InlineData(typeof(NoConstructor))]
    [InlineData(typeof(List<string>))]
    [InlineData(typeof(ConstructorIEnumerableT<string>))]
    [InlineData(typeof(ConstructorICollectionT<int>))]
    [InlineData(typeof(ConstructorICollectionT<string>))]
    [InlineData(typeof(ConstructorIListT<int>))]
    [InlineData(typeof(ConstructorIListT<string>))]
    [InlineData(typeof(AbstractClass))]
    [InlineData(typeof(List<int>))]
    [InlineData(typeof(ConstructorIEnumerableT<int>))]
    [InlineData(typeof(ConstructorICollection))]
    [InlineData(typeof(ConstructorIList))]
    [InlineData(typeof(ConstructorIDictionary))]
    [InlineData(typeof(ArrayList))]
    [InlineData(typeof(Stack))]
    [InlineData(typeof(Queue))]
    [InlineData(typeof(ConstructorIEnumerableT<KeyValuePair<int, int>>))]
    [InlineData(typeof(ConstructorIEnumerableT<KeyValuePair<string, string>>))]
    [InlineData(typeof(ConstructorICollectionT<KeyValuePair<int, int>>))]
    [InlineData(typeof(ConstructorICollectionT<KeyValuePair<string, string>>))]
    [InlineData(typeof(ConstructorIListT<KeyValuePair<string, int>>))]
    [InlineData(typeof(ConstructorIListT<KeyValuePair<int, int>>))]
    [InlineData(typeof(ConstructorIListT<KeyValuePair<string, string>>))]
    [InlineData(typeof(Dictionary<int, int>))]
    [InlineData(typeof(Dictionary<int, string>))]
    [InlineData(typeof(ReadOnlyDictionary<int, int>))]
    [InlineData(typeof(ReadOnlyDictionary<int, string>))]
    [InlineData(typeof(ConstructorIDictionaryT<int, int>))]
    [InlineData(typeof(ConstructorIDictionaryT<int, string>))]
    [InlineData(typeof(ImmutableDictionary<string, int>))]
    [InlineData(typeof(FrozenDictionary<string, int>))]
    [InlineData(typeof(NoConstructorClass))]
    [InlineData(typeof(NoConstructorClass<string, int>))]
    public void Ctor_InvalidDictionaryType_ThrowsArgumentException(Type dictionaryType)
    {
        Assert.Throws<ArgumentException>("dictionaryType", () => new ConstructorDictionaryFactory<string, int>(dictionaryType));
    }

    [Fact]
    public void Begin_End_Success()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Begin_NegativeCount_ThrowsArgumentOutOfRangeException()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        Assert.Throws<ArgumentOutOfRangeException>("count", () => factory.Begin(-1));
    }

    [Fact]
    public void Begin_ThrowingConstructorISetT_Success()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ThrowingConstructorIDictionaryTKeyTValue<string, int>));
        factory.Begin(1);
        Assert.Throws<TargetInvocationException>(() => factory.End());

        // Ensure we can begin again.
        factory.Begin(1);
        Assert.Throws<TargetInvocationException>(() => factory.End());
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));

        // Begin.
        factory.Begin(1);
        factory.Add("key", 1);
        var value = Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key"] = 1 }, value);

        // Begin again.
        factory.Begin(1);
        factory.Add("key", 2);
        value = Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key"] = 2 }, value);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Begin(1);
        factory.Add("key1", 2);

        factory.Add("key2", 3);
        
        var value = Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key1"] = 2, ["key2"] = 3 }, value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", 1));
    }

    [Fact]
    public void Set_Invoke_Success()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Begin(1);
        factory.Add("key1", 1);

        Assert.Equal(new Dictionary<string, int> { ["key1"] = 1 }, Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End()));
    }

    [Fact]
    public void Set_InvokeOutOfRange_Success()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Begin(1);
        factory.Add("key1", 1);
        factory.Add("key2", 2);

        Assert.Equal(new Dictionary<string, int> { ["key1"] = 1, ["key2"] = 2 }, Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End()));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", 1));
    }

    [Fact]
    public void Add_NullKey_ThrowsArgumentNullException()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Begin(1);
        Assert.Throws<ArgumentNullException>("key", () => factory.Add(null!, 1));
    }

    [Fact]
    public void Add_MultipleTimes_ThrowsArgumentException()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Begin(1);
        factory.Add("key", 1);

        Assert.Throws<ArgumentException>(null, () => factory.Add("key", 2));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new ConstructorDictionaryFactory<string, int>(typeof(ReadOnlyDictionary<string, int>));
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<ReadOnlyDictionary<string, int>>(factory.End());
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

    private class ConstructorIDictionary
    {
        public IDictionary Value { get; }

        public ConstructorIDictionary(IDictionary value) => Value = value;
    }

    private class ConstructorIDictionaryT<TKey, TValue> where TKey : notnull
    {
        public IDictionary<TKey, TValue> Value { get; }

        public ConstructorIDictionaryT(IDictionary<TKey, TValue> value) => Value = value;
    }

    private class ThrowingConstructorIDictionaryTKeyTValue<TKey, TValue> where TKey : notnull
    {
        public ThrowingConstructorIDictionaryTKeyTValue(IDictionary<TKey, TValue> value) => throw new NotImplementedException();
    }

    private class NoConstructorClass : IDictionary
    {
        private NoConstructorClass()
        {
        }

        public object? this[object key] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public bool IsFixedSize => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public ICollection Keys => throw new NotImplementedException();

        public ICollection Values => throw new NotImplementedException();

        public int Count => throw new NotImplementedException();

        public bool IsSynchronized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public void Add(object key, object? value) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(object key) => throw new NotImplementedException();

        public void CopyTo(Array array, int index) => throw new NotImplementedException();

        public IDictionaryEnumerator GetEnumerator() => throw new NotImplementedException();

        public void Remove(object key) => throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private class NoConstructorClass<TKey, TValue> : IDictionary<TKey, TValue> where TKey : notnull
    {
        private NoConstructorClass()
        {
        }

        public TValue this[TKey key] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public ICollection<TKey> Keys => throw new NotImplementedException();

        public ICollection<TValue> Values => throw new NotImplementedException();

        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(TKey key, TValue value) => throw new NotImplementedException();

        public void Add(KeyValuePair<TKey, TValue> item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(KeyValuePair<TKey, TValue> item) => throw new NotImplementedException();

        public bool ContainsKey(TKey key) => throw new NotImplementedException();

        public void CopyTo(KeyValuePair<TKey, TValue>[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator<KeyValuePair<TKey, TValue>> GetEnumerator() => throw new NotImplementedException();

        public bool Remove(TKey key) => throw new NotImplementedException();

        public bool Remove(KeyValuePair<TKey, TValue> item) => throw new NotImplementedException();

        public bool TryGetValue(TKey key, [MaybeNullWhen(false)] out TValue value) => throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
