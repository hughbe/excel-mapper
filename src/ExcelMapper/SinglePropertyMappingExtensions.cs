using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Mappers;
using ExcelMapper.Mappings.Readers;
using ExcelMapper.Mappings.Support;
using ExcelMapper.Mappings.Transformers;

namespace ExcelMapper
{
    public delegate T ConvertUsingSimpleMappingDelegate<T>(string stringValue);

    public static class SinglePropertyMappingExtensions
    {
        public static T WithColumnName<T>(this T mapping, string columnName) where T : ISinglePropertyMapping
        {
            return mapping
                .WithReader(new ColumnNameValueReader(columnName));
        }

        public static T WithColumnIndex<T>(this T mapping, int columnIndex) where T : ISinglePropertyMapping
        {
            return mapping
                .WithReader(new ColumnIndexValueReader(columnIndex));
        }

        public static T WithReader<T>(this T mapping, ICellValueReader reader) where T : ISinglePropertyMapping
        {
            if (reader == null)
            {
                throw new ArgumentNullException(nameof(reader));
            }

            if (mapping.CellReader is OptionalCellValueReader optionalMapping)
            {
                optionalMapping.InnerReader = reader;
            }
            else
            {
                mapping.CellReader = reader;
            }

            return mapping;
        }

        public static T MakeOptional<T>(this T mapping) where T : ISinglePropertyMapping
        {
            if (mapping.CellReader is OptionalCellValueReader)
            {
                throw new ExcelMappingException("Mapping is already optional.");
            }

            mapping.CellReader = new OptionalCellValueReader(mapping.CellReader);
            return mapping;
        }

        public static T WithTrim<T>(this T mapping) where T : ISinglePropertyMapping
        {
            var transformer = new TrimCellValueTransformer();
            mapping.AddCellValueTransformer(transformer);
            return mapping;
        }

        public static TMapping WithMappingItems<TMapping>(this TMapping mapping, params ICellValueMapper[] mappings) where TMapping : ISinglePropertyMapping
        {
            if (mappings == null)
            {
                throw new ArgumentNullException(nameof(mappings));
            }

            foreach (ICellValueMapper mappingItem in mappings)
            {
                mapping.AddCellValueMapper(mappingItem);
            }

            return mapping;
        }

        public static TMapping WithMapping<TMapping, T>(this TMapping mapping, IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer = null) where TMapping : ISinglePropertyMapping<T>
        {
            var item = new DictionaryMapper<T>(mappingDictionary, comparer);
            mapping.AddCellValueMapper(item);
            return mapping;
        }

        public static SingleExcelPropertyMap<DateTime> WithDateFormats(this SingleExcelPropertyMap<DateTime> mapping, params string[] formats)
        {
            mapping.AddFormats(formats);
            return mapping;
        }

        public static SingleExcelPropertyMap<DateTime> WithDateFormats(this SingleExcelPropertyMap<DateTime> mapping, IEnumerable<string> formats) 
        {
            return mapping.WithDateFormats(formats?.ToArray());
        }

        public static SingleExcelPropertyMap<DateTime?> WithDateFormats(this SingleExcelPropertyMap<DateTime?> mapping, params string[] formats)
        {
            mapping.AddFormats(formats);
            return mapping;
        }

        public static SingleExcelPropertyMap<DateTime?> WithDateFormats(this SingleExcelPropertyMap<DateTime?> mapping, IEnumerable<string> formats)
        {
            return mapping.WithDateFormats(formats?.ToArray());
        }

        private static void AddFormats(this ISinglePropertyMapping mapping, string[] formats)
        {
            if (formats == null)
            {
                throw new ArgumentNullException(nameof(formats));
            }

            if (formats.Length == 0)
            {
                throw new ArgumentException("Formats cannot be empty.", nameof(formats));
            }

            DateTimeMapper dateTimeItem = (DateTimeMapper)mapping.CellValueMappers.FirstOrDefault(item => item is DateTimeMapper);
            if (dateTimeItem == null)
            {
                dateTimeItem = new DateTimeMapper();
                mapping.AddCellValueMapper(dateTimeItem);
            }

            dateTimeItem.Formats = formats;
        }

        public static TMapping WithConverter<TMapping, T>(this TMapping mapping, ConvertUsingSimpleMappingDelegate<T> converter) where TMapping : ISinglePropertyMapping<T>
        {
            if (converter == null)
            {
                throw new ArgumentNullException(nameof(converter));
            }

            ConvertUsingMappingDelegate actualConverter = (ReadCellValueResult mapResult, ref object value) =>
            {
                try
                {
                    value = converter(mapResult.StringValue);
                    return PropertyMappingResultType.Success;
                }
                catch
                {
                    return PropertyMappingResultType.Invalid;
                }
            };

            var item = new ConvertUsingMapper(actualConverter);
            mapping.AddCellValueMapper(item);
            return mapping;
        }

        public static TMapping WithValueFallback<TMapping>(this TMapping mapping, object defaultValue) where TMapping : ISinglePropertyMapping
        {
            return mapping
                .WithEmptyFallback(defaultValue)
                .WithInvalidFallback(defaultValue);
        }

        public static TMapping WithThrowingFallback<TMapping>(this TMapping mapping) where TMapping : ISinglePropertyMapping
        {
            return mapping
                .WithThrowingEmptyFallback()
                .WithThrowingInvalidFallback();
        }

        public static TMapping WithEmptyFallback<TMapping>(this TMapping mapping, object fallbackValue) where TMapping : ISinglePropertyMapping
        {
            return mapping
                .WithEmptyFallbackItem(new FixedValueFallback(fallbackValue));
        }

        public static TMapping WithEmptyFallbackItem<TMapping>(this TMapping mapping, IFallbackItem fallbackItem) where TMapping : ISinglePropertyMapping
        {
            mapping.EmptyFallback = fallbackItem ?? throw new ArgumentNullException(nameof(fallbackItem));
            return mapping;
        }

        public static TMapping WithThrowingEmptyFallback<TMapping>(this TMapping mapping) where TMapping : ISinglePropertyMapping
        {
            return mapping
                .WithEmptyFallbackItem(new ThrowFallback());
        }

        public static TMapping WithThrowingInvalidFallback<TMapping>(this TMapping mapping) where TMapping : ISinglePropertyMapping
        {
            return mapping
                .WithInvalidFallbackItem(new ThrowFallback());
        }

        public static TMapping WithInvalidFallback<TMapping>(this TMapping mapping, object fallbackValue) where TMapping : ISinglePropertyMapping
        {
            return mapping
                .WithInvalidFallbackItem(new FixedValueFallback(fallbackValue));
        }

        public static TMapping WithInvalidFallbackItem<TMapping>(this TMapping mapping, IFallbackItem fallbackItem) where TMapping : ISinglePropertyMapping
        {
            mapping.InvalidFallback = fallbackItem ?? throw new ArgumentNullException(nameof(fallbackItem));
            return mapping;
        }
    }
}
