using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Mappings;

namespace ExcelMapper
{
    public abstract class EnumerablePropertyMapping<T> : EnumerablePropertyMapping
    {
        public SinglePropertyMapping<T> ElementMapping { get; private set; }

        public IMultiPropertyMapper Mapper { get; internal set; }

        public EnumerablePropertyMapping(MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(member)
        {
            ElementMapping = new SinglePropertyMapping<T>(member, emptyValueStrategy);

            var columnMapper = new ColumnPropertyMapper(member.Name);
            Mapper = new SplitPropertyMapper(columnMapper); 
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
            MapResult[] values = Mapper.GetValues(sheet, rowIndex, reader).ToArray();
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
            var mapper = new ColumnPropertyMapper(columnName);
            if (Mapper is SplitPropertyMapper splitPropertyMapper)
            {
                splitPropertyMapper.Mapper = mapper;
            }
            else
            {
                Mapper = new SplitPropertyMapper(mapper);
            }

            return this;
        }

        public EnumerablePropertyMapping<T> WithIndex(int index)
        {
            var mapper = new IndexPropertyMapper(index);
            if (Mapper is SplitPropertyMapper splitPropertyMapper)
            {
                splitPropertyMapper.Mapper = mapper;
            }
            else
            {
                Mapper = new SplitPropertyMapper(mapper);
            }

            return this;
        }

        public EnumerablePropertyMapping<T> WithSeparators(params char[] separators)
        {
            if (!(Mapper is SplitPropertyMapper splitPropertyMapper))
            {
                throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
            }

            splitPropertyMapper.Separators = separators;
            return this;
        }

        public EnumerablePropertyMapping<T> WithSeparators(IEnumerable<char> separators)
        {
            return WithSeparators(separators?.ToArray());
        }

        public EnumerablePropertyMapping<T> WithColumnNames(params string[] columnNames)
        {
            Mapper = new ColumnsNamesPropertyMapper(columnNames);
            return this;
        }

        public EnumerablePropertyMapping<T> WithColumnNames(IEnumerable<string> columnNames)
        {
            return WithColumnNames(columnNames?.ToArray());
        }

        public EnumerablePropertyMapping<T> WithIndices(params int[] indices)
        {
            Mapper = new ColumnIndicesPropertyMapper(indices);
            return this;
        }

        public EnumerablePropertyMapping<T> WithIndices(IEnumerable<int> indices)
        {
            return WithIndices(indices?.ToArray());
        }
    }
}
