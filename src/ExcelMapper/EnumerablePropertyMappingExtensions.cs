using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Mappings;

namespace ExcelMapper
{
    public static class EnumerablePropertyMappingExtensions
    {
        public static EnumerablePropertyMapping<T> WithColumnName<T>(this EnumerablePropertyMapping<T> mapping, string columnName)
        {
            var mapper = new ColumnPropertyMapper(columnName);
            if (mapping.Mapper is SplitPropertyMapper splitPropertyMapper)
            {
                splitPropertyMapper.Mapper = mapper;
            }
            else
            {
                mapping.Mapper = new SplitPropertyMapper(mapper);
            }

            return mapping;
        }

        public static EnumerablePropertyMapping<T> WithIndex<T>(this EnumerablePropertyMapping<T> mapping, int index)
        {
            var mapper = new IndexPropertyMapper(index);
            if (mapping.Mapper is SplitPropertyMapper splitPropertyMapper)
            {
                splitPropertyMapper.Mapper = mapper;
            }
            else
            {
                mapping.Mapper = new SplitPropertyMapper(mapper);
            }

            return mapping;
        }

        public static EnumerablePropertyMapping<T> WithSeparators<T>(this EnumerablePropertyMapping<T> mapping, params char[] separators)
        {
            return mapping.WithSeparators((IEnumerable<char>)separators);
        }

        public static EnumerablePropertyMapping<T> WithSeparators<T>(this EnumerablePropertyMapping<T> mapping, IEnumerable<char> separators)
        {
            if (!(mapping.Mapper is SplitPropertyMapper splitPropertyMapper))
            {
                throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
            }

            splitPropertyMapper.Separators = separators?.ToArray();
            return mapping;
        }

        public static EnumerablePropertyMapping<T> WithColumnNames<T>(this EnumerablePropertyMapping<T> mapping, params string[] columnNames)
        {
            return mapping.WithColumnNames((IEnumerable<string>)columnNames);
        }

        public static EnumerablePropertyMapping<T> WithColumnNames<T>(this EnumerablePropertyMapping<T> mapping, IEnumerable<string> columnNames)
        {
            mapping.Mapper = new ColumnsPropertyMapper(columnNames);
            return mapping;
        }

        public static EnumerablePropertyMapping<T> WithIndices<T>(this EnumerablePropertyMapping<T> mapping, params int[] indices)
        {
            return mapping.WithIndices((IEnumerable<int>)indices);
        }

        public static EnumerablePropertyMapping<T> WithIndices<T>(this EnumerablePropertyMapping<T> mapping, IEnumerable<int> indices)
        {
            mapping.Mapper = new IndicesPropertyMapper(indices);
            return mapping;
        }
    }
}
