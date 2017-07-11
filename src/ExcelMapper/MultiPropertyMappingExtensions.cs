using System.Collections.Generic;
using ExcelMapper.Mappings;

namespace ExcelMapper
{
    public static class MultiPropertyMappingExtensions
    {
        public static TMapping WithColumnNames<TMapping>(this TMapping mapping, params string[] columnNames) where TMapping : MultiPropertyMapping
        {
            return mapping.WithColumnNames((IEnumerable<string>)columnNames);
        }

        public static TMapping WithColumnNames<TMapping>(this TMapping mapping, IEnumerable<string> columnNames) where TMapping : MultiPropertyMapping
        {
            mapping.Mapper = new ColumnsPropertyMapper(columnNames);
            return mapping;
        }

        public static TMapping WithIndices<TMapping>(this TMapping mapping, params int[] indices) where TMapping : MultiPropertyMapping
        {
            return mapping.WithIndices((IEnumerable<int>)indices);
        }

        public static TMapping WithIndices<TMapping>(this TMapping mapping, IEnumerable<int> indices) where TMapping : MultiPropertyMapping
        {
            mapping.Mapper = new IndicesPropertyMapper(indices);
            return mapping;
        }
    }
}
