using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Readers;

namespace ExcelMapper
{
    public abstract class EnumerablePropertyMapping<T> : EnumerablePropertyMapping
    {
        public SinglePropertyMapping<T> ElementMapping { get; private set; }

        public IMultipleValuesReader ColumnsReader { get; internal set; }

        public EnumerablePropertyMapping(MemberInfo member, SinglePropertyMapping<T> elementMapping) : base(member)
        {
            ElementMapping = elementMapping ?? throw new ArgumentNullException(nameof(elementMapping));

            var columnReader = new ColumnNameReader(member.Name);
            ColumnsReader = new SplitColumnReader(columnReader); 
        }

        public EnumerablePropertyMapping<T> WithElementMapping(Func<SinglePropertyMapping<T>, SinglePropertyMapping<T>> elementMapping)
        {
            if (elementMapping == null)
            {
                throw new ArgumentNullException(nameof(elementMapping));
            }

            ElementMapping = elementMapping(ElementMapping) ?? throw new ArgumentNullException(nameof(elementMapping));
            return this;
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            ReadResult[] values = ColumnsReader.GetValues(sheet, rowIndex, reader).ToArray();
            var elements = new List<T>(values.Length);

            for (int i = 0; i < values.Length; i++)
            {
                T value = (T)ElementMapping.GetPropertyValue(sheet, rowIndex, reader, values[i]);
                elements.Add(value);
            }

            return CreateFromElements(elements);
        }

        public abstract object CreateFromElements(IEnumerable<T> elements);

        public EnumerablePropertyMapping<T> WithColumnName(string columnName)
        {
            var columnReader = new ColumnNameReader(columnName);
            if (ColumnsReader is SplitColumnReader splitColumnReader)
            {
                splitColumnReader.ColumnReader = columnReader;
            }
            else
            {
                ColumnsReader = new SplitColumnReader(columnReader);
            }

            return this;
        }

        public EnumerablePropertyMapping<T> WithColumnIndex(int index)
        {
            var reader = new ColumnIndexReader(index);
            if (ColumnsReader is SplitColumnReader splitColumnReader)
            {
                splitColumnReader.ColumnReader = reader;
            }
            else
            {
                ColumnsReader = new SplitColumnReader(reader);
            }

            return this;
        }

        public EnumerablePropertyMapping<T> WithSeparators(params char[] separators)
        {
            if (!(ColumnsReader is SplitColumnReader splitColumnReader))
            {
                throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
            }

            splitColumnReader.Separators = separators;
            return this;
        }

        public EnumerablePropertyMapping<T> WithSeparators(IEnumerable<char> separators)
        {
            return WithSeparators(separators?.ToArray());
        }

        public EnumerablePropertyMapping<T> WithColumnNames(params string[] columnNames)
        {
            ColumnsReader = new MultipleColumnNamesReader(columnNames);
            return this;
        }

        public EnumerablePropertyMapping<T> WithColumnNames(IEnumerable<string> columnNames)
        {
            return WithColumnNames(columnNames?.ToArray());
        }

        public EnumerablePropertyMapping<T> WithColumnIndices(params int[] indices)
        {
            ColumnsReader = new MultipleColumnIndicesReader(indices);
            return this;
        }

        public EnumerablePropertyMapping<T> WithColumnIndices(IEnumerable<int> indices)
        {
            return WithColumnIndices(indices?.ToArray());
        }
    }
}
