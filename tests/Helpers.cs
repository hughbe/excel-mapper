using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text;

namespace ExcelMapper.Tests;

public static class Helpers
{
    public static bool Initialized { get; private set; }

    public static ExcelImporter GetImporter(string name) => new(GetResource(name));

    public static string GetResourcePath(string name) => Path.GetFullPath(Path.Combine("Resources", name));

    public static Stream GetResource(string name)
    {
        if (!Initialized)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Initialized = true;
        }

        return File.OpenRead(GetResourcePath(name));
    }

    public class TestClass
    {
        public string Value { get; set; } = default!;
        public object ObjectValue { get; set; } = default!;
        public NestedClass NestedValue { get; set; } = default!;
        public DateTime DateTimeValue { get; set; }
        public DateTime? NullableDateTimeValue { get; set; }
        public DateTimeOffset DateTimeOffsetValue { get; set; }
        public DateTimeOffset? NullableDateTimeOffsetValue { get; set; }
        public TimeSpan TimeSpanValue { get; set; }
        public TimeSpan? NullableTimeSpanValue { get; set; }
        public DateOnly DateOnlyValue { get; set; }
        public DateOnly? NullableDateOnlyValue { get; set; }
        public TimeOnly TimeOnlyValue { get; set; }
        public TimeOnly? NullableTimeOnlyValue { get; set; }
        public IListInterface UnknownInterfaceValue { get; set; } = default!;
        public ConcreteIEnumerable ConcreteIEnumerable { get; set; } = default!;
        public ConcreteIDictionary ConcreteIDictionary { get; set; } = default!;
        public ImmutableDictionary<string, string>.Builder IDictionaryNoConstructor { get; set; } = default!;
        public IList<IList<string>> CantMapElementType { get; set; } = default!;
        public IDictionary<string, IList<string>> CantMapDictionaryValueType { get; set; } = default!;
        public string[,] MultiDimensionalArray { get; set; } = default!;

        public InvalidIListMemberType InvalidIListMemberType { get; set; } = default!;
        public InvalidIDictionaryMemberType InvalidIDictionaryMemberType { get; set; } = default!;

        public event EventHandler Event { add { } remove { } }

        public class NestedClass
        {
            public int IntValue { get; set; }
        }
    }

    public class InvalidIListMemberType
    {
        public IListInterface ConcreteIEnumerable { get; set; } = default!;
    }

    public class InvalidIDictionaryMemberType
    {
        public IDictionaryInterface ConcreteIEnumerable { get; set; } = default!;
    }

    public interface IEmptyInterface { }

    public interface IListInterface : IEmptyInterface, IList<string>
    {
    }

    public interface IDictionaryInterface : IEmptyInterface, IDictionary<string, string>
    {
    }

    public class ConcreteIEnumerable : IEmptyInterface, IEnumerable<string>
    {
        public IEnumerator<string> GetEnumerator() => throw new NotImplementedException();
        IEnumerator IEnumerable.GetEnumerator() => throw new NotImplementedException();
    }

    public class ConcreteIDictionary : IEmptyInterface, IDictionary<string, string>
    {
        public string this[string key] { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public ICollection<string> Keys => throw new NotImplementedException();

        public ICollection<string> Values => throw new NotImplementedException();

        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(string key, string value) => throw new NotImplementedException();

        public void Add(KeyValuePair<string, string> item) => throw new NotImplementedException();

        public void Clear() => throw new NotImplementedException();

        public bool Contains(KeyValuePair<string, string> item) => throw new NotImplementedException();

        public bool ContainsKey(string key) => throw new NotImplementedException();

        public void CopyTo(KeyValuePair<string, string>[] array, int arrayIndex) => throw new NotImplementedException();

        public IEnumerator<KeyValuePair<string, string>> GetEnumerator() => throw new NotImplementedException();

        public bool Remove(string key) => throw new NotImplementedException();

        public bool Remove(KeyValuePair<string, string> item) => throw new NotImplementedException();

        public bool TryGetValue(string key, [NotNullWhen(true)] out string? value) => throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => throw new NotImplementedException();
    }
}
