using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelMapper.Tests
{
    public static class Helpers
    {
        public static bool Initialized { get; private set; }

        public static ExcelImporter GetImporter(string name) => new ExcelImporter(GetResource(name));

        public static Stream GetResource(string name)
        {
            if (!Initialized)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                Initialized = true;
            }

            string filePath = Path.GetFullPath(Path.Combine("Resources", name));
            return File.OpenRead(filePath);
        }

        public class TestClass
        {
            public string Value { get; set; }
            public object ObjectValue { get; set; }
            public NestedClass NestedValue { get; set; }
            public DateTime DateValue { get; set; }
            public DateTime? NullableDateValue { get; set; }
            public IListInterface UnknownInterfaceValue { get; set; }
            public ConcreteIEnumerable ConcreteIEnumerable { get; set; }
            public ConcreteIDictionary ConcreteIDictionary { get; set; }
            public IList<IList<string>> CantMapElementType { get; set; }
            public IDictionary<string, IList<string>> CantMapDictionaryValueType { get; set; }
            public string[,] MultiDimensionalArray { get; set; }

            public InvalidIListMemberType InvalidIListMemberType { get; set; }
            public InvalidIDictionaryMemberType InvalidIDictionaryMemberType { get; set; }

            public event EventHandler Event { add { } remove { } }

            public class NestedClass
            {
                public int IntValue { get; set; }
            }
        }

        public class InvalidIListMemberType
        {
            public IListInterface ConcreteIEnumerable { get; set; }
        }

        public class InvalidIDictionaryMemberType
        {
            public IDictionaryInterface ConcreteIEnumerable { get; set; }
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

            public bool TryGetValue(string key, out string value) => throw new NotImplementedException();

            IEnumerator IEnumerable.GetEnumerator() => throw new NotImplementedException();
        }
    }
}
