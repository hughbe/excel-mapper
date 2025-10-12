using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using Xunit;

namespace ExcelMapper.Factories;

public class IDictionaryTImplementingFactoryTests
{
    [Theory]
    [InlineData(typeof(Dictionary<string, int>))]
    [InlineData(typeof(IDictionaryGeneric<string, int>))]
    public void Ctor_Type(Type dictionaryType)
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(dictionaryType);
        Assert.Equal(dictionaryType, factory.DictionaryType);
    }

    [Fact]
    public void Ctor_NullDictionaryType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("dictionaryType", () => new IDictionaryTImplementingFactory<string, int>(null!));
    }

    [Theory]
    [InlineData(typeof(IEnumerable))]
    [InlineData(typeof(IEnumerable<int>))]
    [InlineData(typeof(IEnumerable<KeyValuePair<string, int>>))]
    [InlineData(typeof(ICollection))]
    [InlineData(typeof(ICollection<int>))]
    [InlineData(typeof(ICollection<KeyValuePair<string, int>>))]
    [InlineData(typeof(IList))]
    [InlineData(typeof(IList<int>))]
    [InlineData(typeof(IList<KeyValuePair<string, int>>))]
    [InlineData(typeof(IDictionary))]
    [InlineData(typeof(IDictionary<string, int>))]
    [InlineData(typeof(IReadOnlyDictionary<string, int>))]
    [InlineData(typeof(int[]))]
    [InlineData(typeof(ArrayList))]
    [InlineData(typeof(Stack))]
    [InlineData(typeof(Queue))]
    [InlineData(typeof(Dictionary<string, string>))]
    [InlineData(typeof(IDictionaryNonGeneric))]
    [InlineData(typeof(IDictionaryGeneric<int, string>))]
    [InlineData(typeof(AbstractClass))]
    public void Ctor_InvalidDictionaryType_ThrowsArgumentException(Type dictionaryType)
    {
        Assert.Throws<ArgumentException>("dictionaryType", () => new IDictionaryTImplementingFactory<string, int>(dictionaryType));
    }

    [Fact]
    public void Begin_End_Success()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));

        // Begin.
        factory.Begin(1);
        factory.Add("key", 1);
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key"] = 1 }, value);

        // Begin again.
        factory.Begin(1);
        factory.Add("key", 2);
        value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key"] = 2 }, value);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Begin(1);
        factory.Add("key1", 2);

        factory.Add("key2", 3);

        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key1"] = 2, ["key2"] = 3 }, value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", 1));
    }

    [Fact]
    public void Set_Invoke_Success()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Begin(1);
        factory.Add("key1", 1);

        Assert.Equal(new Dictionary<string, int> { ["key1"] = 1 }, Assert.IsType<Dictionary<string, int>>(factory.End()));
    }

    [Fact]
    public void Set_InvokeOutOfRange_Success()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Begin(1);
        factory.Add("key1", 1);
        factory.Add("key2", 2);

        Assert.Equal(new Dictionary<string, int> { ["key1"] = 1, ["key2"] = 2 }, Assert.IsType<Dictionary<string, int>>(factory.End()));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", 1));
    }

    [Fact]
    public void Set_NullKey_ThrowsArgumentNullException()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Begin(1);
        Assert.Throws<ArgumentNullException>("key", () => factory.Add(null!, 1));
    }

    [Fact]
    public void Set_MultipleTimes_ThrowsArgumentException()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Begin(1);
        factory.Add("key", 1);

        Assert.Throws<ArgumentException>(null, () => factory.Add("key", 2));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new IDictionaryTImplementingFactory<string, int>(typeof(Dictionary<string, int>));
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    private abstract class AbstractClass
    {
    }

    private class IDictionaryNonGeneric : IDictionary
    {
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

        IEnumerator IEnumerable.GetEnumerator() => throw new NotImplementedException();
    }

    private class IDictionaryGeneric<TKey, TValue> : IDictionary<TKey, TValue> where TKey : notnull
    {
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

        IEnumerator IEnumerable.GetEnumerator() => throw new NotImplementedException();
    }
}
