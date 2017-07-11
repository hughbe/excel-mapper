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
            public DateTime DateValue { get; set; }
            public IListInterface UnknownInterfaceValue { get; set; }
            public ConcreteIEnumerable ConcreteIEnumerable { get; set; }

            public event EventHandler Event { add { } remove { } }
        }

        public interface IListInterface : IList<string>
        {
        }

        public class ConcreteIEnumerable : IEnumerable<string>
        {
            public IEnumerator<string> GetEnumerator() => throw new NotImplementedException();
            IEnumerator IEnumerable.GetEnumerator() => throw new NotImplementedException();
        }
    }
}
