using System;
using System.Collections;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics.CodeAnalysis;
using Xunit;

namespace ExcelMapper.Factories;

public class AddDictionaryFactoryTests
{
    [Theory]
    [InlineData(typeof(Dictionary<string, string>))]
    [InlineData(typeof(AddClass<string, string>))]
    [InlineData(typeof(Hashtable))]
    [InlineData(typeof(StringDictionary))]
    public void Ctor_Type(Type dictionaryType)
    {
        var factory = new AddDictionaryFactory<string, string>(dictionaryType);
        Assert.Equal(dictionaryType, factory.DictionaryType);
    }

    [Fact]
    public void Ctor_NullDictionaryType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("dictionaryType", () => new AddDictionaryFactory<string, string>(null!));
    }

    [Theory]
    [InlineData(typeof(IEnumerable))]
    [InlineData(typeof(IEnumerable<int>))]
    [InlineData(typeof(ICollection))]
    [InlineData(typeof(ICollection<int>))]
    [InlineData(typeof(IList))]
    [InlineData(typeof(IList<int>))]
    [InlineData(typeof(IDictionary))]
    [InlineData(typeof(IDictionary<string, string>))]
    [InlineData(typeof(int[]))]
    [InlineData(typeof(ReadOnlyDictionary<string, string>))]
    [InlineData(typeof(Dictionary<int, string>))]
    [InlineData(typeof(Dictionary<string, int>))]
    [InlineData(typeof(AddClass<int, string>))]
    [InlineData(typeof(AddClass<string, int>))]
    [InlineData(typeof(AddOneClass))]
    [InlineData(typeof(AddThreeClass))]
    [InlineData(typeof(NonEnumerableAddClass<string, string>))]
    [InlineData(typeof(FrozenDictionary<string, string>))]
    [InlineData(typeof(ImmutableDictionary<string, string>))]
    [InlineData(typeof(ReadOnlyDictionary<string, int>))]
    [InlineData(typeof(NoConstructorClass))]
    [InlineData(typeof(NoConstructorClass<string, int>))]
    public void Ctor_InvalidDictionaryType_ThrowsArgumentException(Type dictionaryType)
    {
        Assert.Throws<ArgumentException>("dictionaryType", () => new AddDictionaryFactory<string, string>(dictionaryType));
    }
    
    [Fact]
    public void Begin_End_Success()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Begin_NegativeCount_ThrowsArgumentOutOfRangeException()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        Assert.Throws<ArgumentOutOfRangeException>("count", () => factory.Begin(-1));
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));

        // Begin.
        factory.Begin(1);
        factory.Add("key", "1");
        var value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Single(value);
        Assert.Equal("1", value["key"]);

        // Begin again.
        factory.Begin(1);
        factory.Add("key", "2");
        value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Single(value);
        Assert.Equal("2", value["key"]);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Begin(1);
        factory.Add("key1", "2");

        factory.Add("key2", "3");
        
        var value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Equal(2, value.Count);
        Assert.Equal("2", value["key1"]);
        Assert.Equal("3", value["key2"]);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", "1"));
    }

    [Fact]
    public void Set_Invoke_Success()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Begin(1);
        factory.Add("key1", "1");

        var value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Single(value);
        Assert.Equal("1", value["key1"]);
    }

    [Fact]
    public void Set_InvokeOutOfRange_Success()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Begin(1);
        factory.Add("key1", "1");
        factory.Add("key2", "2");

        var value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Equal(2, value.Count);
        Assert.Equal("1", value["key1"]);
        Assert.Equal("2", value["key2"]);
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", "1"));
    }

    [Fact]
    public void Add_NullKey_ThrowsArgumentNullException()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Begin(1);
        Assert.Throws<ArgumentNullException>("key", () => factory.Add(null!, "1"));
    }

    [Fact]
    public void Add_MultipleTimes_ThrowsArgumentException()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Begin(1);
        factory.Add("key", "1");

        Assert.Throws<ArgumentException>(null, () => factory.Add("key", "2"));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new AddDictionaryFactory<string, string>(typeof(StringDictionary));
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<StringDictionary>(factory.End());
        Assert.Equal([], value);
    }

    private class AddOneClass : IEnumerable
    {
        public void Add(string item)
        {
        }

        public IEnumerator GetEnumerator() => throw new NotImplementedException();
    }

    private class AddClass<TKey, TValue> : IEnumerable where TKey : notnull
    {
        public void Add(TKey key, TValue value)
        {
        }

        public IEnumerator GetEnumerator() => throw new NotImplementedException();
    }

    private class NonEnumerableAddClass<TKey, TValue> where TKey : notnull
    {
        public void Add(TKey key, TValue value) => throw new NotImplementedException();
    }

    private class AddThreeClass : IEnumerable
    {
        public void Add(string item, string item2, string item3)
        {
        }

        public IEnumerator GetEnumerator() => throw new NotImplementedException();
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
