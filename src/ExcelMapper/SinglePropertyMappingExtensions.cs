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
                .WithReader(new ColumnNameReader(columnName));
        }

        public static T WithColumnIndex<T>(this T mapping, int columnIndex) where T : ISinglePropertyMapping
        {
            return mapping
                .WithReader(new ColumnIndexReader(columnIndex));
        }

        public static T WithReader<T>(this T mapping, ISingleValueReader reader) where T : ISinglePropertyMapping
        {
            if (reader == null)
            {
                throw new ArgumentNullException(nameof(reader));
            }

            if (mapping.Reader is OptionalColumnReader optionalMapping)
            {
                optionalMapping.InnerReader = reader;
            }
            else
            {
                mapping.Reader = reader;
            }

            return mapping;
        }

        public static T MakeOptional<T>(this T mapping) where T : ISinglePropertyMapping
        {
            if (mapping.Reader is OptionalColumnReader)
            {
                throw new ExcelMappingException("Mapping is already optional.");
            }

            mapping.Reader = new OptionalColumnReader(mapping.Reader);
            return mapping;
        }

        public static T WithTrim<T>(this T mapping) where T : ISinglePropertyMapping
        {
            var transformer = new TrimStringTransformer();
            mapping.AddStringValueTransformer(transformer);
            return mapping;
        }

        public static TMapping WithMappingItems<TMapping>(this TMapping mapping, params IStringValueMapper[] mappings) where TMapping : ISinglePropertyMapping
        {
            if (mappings == null)
            {
                throw new ArgumentNullException(nameof(mappings));
            }

            foreach (IStringValueMapper mappingItem in mappings)
            {
                mapping.AddMappingItem(mappingItem);
            }

            return mapping;
        }

        public static TMapping WithMapping<TMapping, T>(this TMapping mapping, IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer = null) where TMapping : ISinglePropertyMapping<T>
        {
            var item = new DictionaryMapper<T>(mappingDictionary, comparer);
            mapping.AddMappingItem(item);
            return mapping;
        }

        public static SinglePropertyMapping<DateTime> WithDateFormats(this SinglePropertyMapping<DateTime> mapping, params string[] formats)
        {
            mapping.AddFormats(formats);
            return mapping;
        }

        public static SinglePropertyMapping<DateTime> WithDateFormats(this SinglePropertyMapping<DateTime> mapping, IEnumerable<string> formats) 
        {
            return mapping.WithDateFormats(formats?.ToArray());
        }

        public static SinglePropertyMapping<DateTime?> WithDateFormats(this SinglePropertyMapping<DateTime?> mapping, params string[] formats)
        {
            mapping.AddFormats(formats);
            return mapping;
        }

        public static SinglePropertyMapping<DateTime?> WithDateFormats(this SinglePropertyMapping<DateTime?> mapping, IEnumerable<string> formats)
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

            DateTimeMapper dateTimeItem = (DateTimeMapper)mapping.MappingItems.FirstOrDefault(item => item is DateTimeMapper);
            if (dateTimeItem == null)
            {
                dateTimeItem = new DateTimeMapper();
                mapping.AddMappingItem(dateTimeItem);
            }

            dateTimeItem.Formats = formats;
        }

        public static TMapping WithConverter<TMapping, T>(this TMapping mapping, ConvertUsingSimpleMappingDelegate<T> converter) where TMapping : ISinglePropertyMapping<T>
        {
            if (converter == null)
            {
                throw new ArgumentNullException(nameof(converter));
            }

            ConvertUsingMappingDelegate actualConverter = (ReadResult mapResult, ref object value) =>
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
            mapping.AddMappingItem(item);
            return mapping;
        }

        public static TMapping WithValueFallback<TMapping, T>(this TMapping mapping, T defaultValue) where TMapping : ISinglePropertyMapping<T>
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

        public static TMapping WithEmptyFallback<TMapping, T>(this TMapping mapping, T fallbackValue) where TMapping : ISinglePropertyMapping<T>
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

        public static TMapping WithInvalidFallback<TMapping, T>(this TMapping mapping, T fallbackValue) where TMapping : ISinglePropertyMapping<T>
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
