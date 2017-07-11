using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    public delegate object CreateInstance();
    public delegate PropertyMappingResult ConvertStringValue(string stringValue);
    public delegate void AddElement(object instance, object value);
    public delegate object ReturnInstance(object instance);

    internal class SplitWithDelimiterMappingItem : ISinglePropertyMappingItem
    {
        private static char[] s_defaultDelimiter = new char[] { ',' };

        public char[] Delimiters { get; internal set; }
        public StringSplitOptions Options { get; private set; }

        public CreateInstance CreateDelegate { get; }
        public ConvertStringValue ConvertDelegate { get; }
        public AddElement AddDelegate { get; }
        public ReturnInstance ReturnDelegate { get; }

        public SplitWithDelimiterMappingItem(CreateInstance createDelegate, ConvertStringValue convertDelegate, AddElement addDelegate, ReturnInstance returnDelegate) : this(createDelegate, convertDelegate, addDelegate, returnDelegate, s_defaultDelimiter) { }

        public SplitWithDelimiterMappingItem(CreateInstance createDelegate, ConvertStringValue convertDelegate, AddElement addDelegate, ReturnInstance returnDelegate, IEnumerable<char> delimiters)
        {
            if (delimiters == null)
            {
                throw new ArgumentNullException(nameof(delimiters));
            }

            Delimiters = delimiters.ToArray();

            CreateDelegate = createDelegate;
            ConvertDelegate = convertDelegate;
            AddDelegate = addDelegate;
            ReturnDelegate = returnDelegate;
        }

        public SplitWithDelimiterMappingItem WithOptions(StringSplitOptions options)
        {
            Options = options;
            return this;
        }

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue)
        {
            object instance = CreateDelegate();

            string[] results = stringValue.Split(Delimiters, Options);
            foreach (string result in results)
            {
                PropertyMappingResult elementResult = ConvertDelegate(result);
                if (elementResult.Type == PropertyMappingResultType.Invalid)
                {
                    return elementResult;
                }

                AddDelegate(instance, elementResult.Value);
            }

            object convertedInstance = ReturnDelegate(instance);
            return PropertyMappingResult.Success(convertedInstance);
        }
    }
}
