using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Items;
using ExcelMapper.Mappings.Support;
using ExcelMapper.Mappings.Transformers;

namespace ExcelMapper
{
    public delegate T ConvertUsingSimpleMappingDelegate<T>(string stringValue);

    public static class SinglePropertyMappingExtensions
    {
        public static T WithColumnName<T>(this T mapping, string columnName) where T : ISinglePropertyMapping
        {
            mapping.Mapper = new ColumnPropertyMapper(columnName);
            return mapping;
        }

        public static T WithIndex<T>(this T mapping, int index) where T : ISinglePropertyMapping
        {
            mapping.Mapper = new IndexPropertyMapper(index);
            return mapping;
        }

        public static T WithTrim<T>(this T mapping) where T : ISinglePropertyMapping
        {
            var transformer = new TrimStringTransformer();
            mapping.AddStringValueTransformer(transformer);
            return mapping;
        }

        public static TMapping WithMapping<TMapping, T>(this TMapping mapping, IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer = null) where TMapping : ISinglePropertyMapping<T>
        {
            var item = new MapStringValueMappingItem<T>(mappingDictionary, comparer);
            mapping.AddMappingItem(item);
            return mapping;
        }

        public static T WithAdditionalDateFormats<T>(this T mapping, params string[] formats) where T : SinglePropertyMapping<DateTime>
        {
            return mapping.WithAdditionalDateFormats((IEnumerable<string>)formats);
        }

        public static T WithAdditionalDateFormats<T>(this T mapping, IEnumerable<string> formats) where T : SinglePropertyMapping<DateTime>
        {
            if (formats == null)
            {
                throw new ArgumentNullException(nameof(formats));
            }

            ParseAsDateTimeMappingItem dateTimeItem = (ParseAsDateTimeMappingItem)mapping.MappingItems.FirstOrDefault(item => item is ParseAsDateTimeMappingItem);
            if (dateTimeItem == null)
            {
                dateTimeItem = new ParseAsDateTimeMappingItem();
                mapping.AddMappingItem(dateTimeItem);
            }

            dateTimeItem.Formats = dateTimeItem.Formats.Concat(formats).ToArray();
            return mapping;
        }

        public static T WithDelimiters<T>(this T pipeline, params char[] delimiters) where T : ISinglePropertyMapping<IEnumerable>
        {
            return pipeline
                .WithDelimiters((IEnumerable<char>)delimiters);
        }

        public static T WithDelimiters<T>(this T mapping, IEnumerable<char> delimiters) where T : ISinglePropertyMapping<IEnumerable>
        {
            if (delimiters == null)
            {
                throw new ArgumentNullException(nameof(delimiters));
            }

            SplitWithDelimiterMappingItem splitItem = (SplitWithDelimiterMappingItem)mapping.MappingItems.FirstOrDefault(item => item is SplitWithDelimiterMappingItem);
            if (splitItem == null)
            {
                // TODO.
                splitItem = null;
                mapping.AddMappingItem(splitItem);
            }

            splitItem.Delimiters = delimiters.ToArray();
            return mapping;
        }

        public static TMapping WithConverter<TMapping, T>(this TMapping mapping, ConvertUsingSimpleMappingDelegate<T> converter) where TMapping : ISinglePropertyMapping<T>
        {
            if (converter == null)
            {
                throw new ArgumentNullException(nameof(converter));
            }

            ConvertUsingMappingDelegate actualConverter = (stringValue) =>
            {
                try
                {
                    T value = converter(stringValue);
                    return PropertyMappingResult.Success(value);
                }
                catch
                {
                    return PropertyMappingResult.Invalid();
                }
            };

            var item = new ConvertUsingMappingItem(actualConverter);
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

        public static TMapping WithEmptyFallbackItem<TMapping>(this TMapping mapping, ISinglePropertyMappingItem fallbackItem) where TMapping : ISinglePropertyMapping
        {
            mapping.EmptyFallback = fallbackItem;
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

        public static TMapping WithInvalidFallbackItem<TMapping>(this TMapping mapping, ISinglePropertyMappingItem fallbackItem) where TMapping : ISinglePropertyMapping
        {
            mapping.InvalidFallback = fallbackItem;
            return mapping;
        }
    }
}
