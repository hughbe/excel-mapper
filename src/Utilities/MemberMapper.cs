using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Reflection;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;
using ExcelMapper.Utilities;

namespace ExcelMapper.Utilities;

internal static class MemberMapper
{
    internal static ICellReaderFactory GetDefaultCellReaderFactory(MemberInfo member)
    {
        var columnNameAttributes = MemberAttributes.GetMany<ExcelColumnNameAttribute>(member);
        if (columnNameAttributes.Length == 1)
        {
            return new ColumnNameReaderFactory(columnNameAttributes[0].Name, columnNameAttributes[0].Comparison);
        }

        if (columnNameAttributes.Length > 1)
        {
            var factories = new ICellReaderFactory[columnNameAttributes.Length];
            for (int i = 0; i < factories.Length; i++)
            {
                factories[i] = new ColumnNameReaderFactory(columnNameAttributes[i].Name, columnNameAttributes[i].Comparison);
            }

            return new CompositeCellsReaderFactory(factories);
        }

        var columnNamesAttribute = MemberAttributes.Get<ExcelColumnNamesAttribute>(member);
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesReaderFactory([.. columnNamesAttribute.Names], columnNamesAttribute.Comparison);
        }

        // A single [ExcelColumnNameMatching] attributes still represents one column, but multiple options.
        var columnNameMatchingAttribute = MemberAttributes.Get<ExcelColumnMatchingAttribute>(member);
        if (columnNameMatchingAttribute != null)
        {
            var matcher = Activator.CreateInstance(columnNameMatchingAttribute.Type, columnNameMatchingAttribute.ConstructorArguments)!;
            return matcher is IColumnSelector selector
                ? new ColumnSelectorReaderFactory(selector)
                : new ColumnsMatchingReaderFactory((IExcelColumnMatcher)matcher);
        }

        var columnIndexAttributes = MemberAttributes.GetMany<ExcelColumnIndexAttribute>(member);
        if (columnIndexAttributes.Length == 1)
        {
            return new ColumnIndexReaderFactory(columnIndexAttributes[0].Index);
        }

        if (columnIndexAttributes.Length > 1)
        {
            var indices = new int[columnIndexAttributes.Length];
            for (int i = 0; i < indices.Length; i++)
            {
                indices[i] = columnIndexAttributes[i].Index;
            }

            return new ColumnIndicesReaderFactory(indices);
        }

        // [ExcelColumnIndices] attributes still represents one column, but multiple options.
        var columnIndicesAttribute = MemberAttributes.Get<ExcelColumnIndicesAttribute>(member);
        if (columnIndicesAttribute != null)
        {
            return new ColumnIndicesReaderFactory([.. columnIndicesAttribute.Indices]);
        }

        return new ColumnNameReaderFactory(member.Name);
    }

    internal static ICellsReaderFactory GetDefaultSplitCellsReaderFactory(MemberInfo? member, ICellReaderFactory innerReaderFactory)
    {
        if (member != null)
        {
            // If the member has a ExcelSeparator attribute, use that separator.
            if (MemberAttributes.Get<ExcelSeparatorsAttribute>(member) is { } separatorsAttribute)
            {
                if (separatorsAttribute.StringSeparators is { } stringSeparators)
                {
                    return new StringSplitReaderFactory(innerReaderFactory)
                    {
                        Separators = [.. stringSeparators],
                        Options = separatorsAttribute.Options
                    };
                }

                return new CharSplitReaderFactory(innerReaderFactory)
                {
                    Separators = [.. separatorsAttribute.CharSeparators!],
                    Options = separatorsAttribute.Options
                };
            }
        }

        return new CharSplitReaderFactory(innerReaderFactory);
    }
    internal static ICellsReaderFactory? GetDefaultCellsReaderFactory(MemberInfo? member)
    {
        // If no member was specified, read all the cells.
        if (member == null)
        {
            return new AllColumnNamesReaderFactory();
        }

        // [ExcelColumnNames] attributes represent multiple columns.
        var columnNamesAttribute = MemberAttributes.Get<ExcelColumnNamesAttribute>(member);
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesReaderFactory([.. columnNamesAttribute.Names], columnNamesAttribute.Comparison);
        }

        // [ExcelColumnsMatchingAttribute] attributes represent multiple columns.
        var columnNameMatchingAttribute = MemberAttributes.Get<ExcelColumnsMatchingAttribute>(member);
        if (columnNameMatchingAttribute != null)
        {
            var matcher = Activator.CreateInstance(columnNameMatchingAttribute.Type, columnNameMatchingAttribute.ConstructorArguments)!;
            return matcher is IColumnSelector selector
                ? new ColumnSelectorReaderFactory(selector)
                : new ColumnsMatchingReaderFactory((IExcelColumnMatcher)matcher);
        }

        // [ExcelColumnIndices] attributes represents multiple columns.
        var columnIndicesAttribute = MemberAttributes.Get<ExcelColumnIndicesAttribute>(member);
        if (columnIndicesAttribute != null)
        {
            return new ColumnIndicesReaderFactory([.. columnIndicesAttribute.Indices]);
        }

        return null;
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    internal static bool AddMappers<T>(IValuePipeline pipeline, MemberInfo member)
    {
        var addDefaultMappers = true;

        // If the member has ExcelMapper attributes, add the mappers.
        var mapperAttributes = MemberAttributes.GetMany<ExcelMapperAttribute>(member);
        bool hasMapperAttributes = false;
        foreach (var mapperAttribute in mapperAttributes)
        {
            hasMapperAttributes = true;
            var mapper = (ICellMapper)Activator.CreateInstance(mapperAttribute.Type, mapperAttribute.ConstructorArguments)!;
            pipeline.Mappers.Add(mapper);
        }

        // If the member has ExcelConverter attributes, add their converters.
        foreach (var converterAttribute in MemberAttributes.GetMany<ExcelConverterAttribute>(member))
        {
            var converter = Activator.CreateInstance(converterAttribute.Type, converterAttribute.ConstructorArguments)!;
            if (CreateConverterMapper<T>(converter, member) is { } converterMapper)
            {
                hasMapperAttributes = true;
                pipeline.Mappers.Add(converterMapper);
            }
        }

        // Since explicit mappers were added, do not add default mappers.
        if (hasMapperAttributes)
        {
            addDefaultMappers = false;
        }

        // If the member has any ExcelMappingDictionary attributes, add a MappingDictionaryMapper.
        var mappingDictionaryAttributes = MemberAttributes.GetMany<ExcelMappingDictionaryAttribute>(member);
        Dictionary<string, object?>? mappingDictionary = null;
        foreach (var mappingDictionaryAttribute in mappingDictionaryAttributes)
        {
            mappingDictionary ??= new Dictionary<string, object?>();
            mappingDictionary.Add(mappingDictionaryAttribute.Value, mappingDictionaryAttribute.MappedValue);
        }

        if (mappingDictionary != null)
        {
            // If the member has a ExcelMappingDictionaryComparer attribute, get the comparer.
            IEqualityComparer<string>? comparer = null;
            if (MemberAttributes.Get<ExcelMappingDictionaryComparerAttribute>(member) is { } comparerAttribute)
            {
                comparer = StringComparerUtilities.FromComparison(comparerAttribute.Comparison);
            }

            // If the member has a ExcelMappingDictionaryBehavior attribute, get the behavior.
            var behavior = MappingDictionaryMapperBehavior.Optional;
            if (MemberAttributes.Get<ExcelMappingDictionaryBehaviorAttribute>(member) is { } behaviorAttribute)
            {
                behavior = behaviorAttribute.Behavior;
            }

            pipeline.Mappers.Add(new MappingDictionaryMapper<object?>(mappingDictionary, comparer, behavior));

            // If the dictionary mapper was added as required, do not add default mappers.
            if (behavior == MappingDictionaryMapperBehavior.Required)
            {
                addDefaultMappers = false;
            }
        }

        return addDefaultMappers;
    }

    /// <summary>
    /// Creates the mapper for a converter given by <see cref="ExcelConverterAttribute"/>: an
    /// <see cref="ICellConverter{T}"/>, or for a <c>T?</c> member, an <see cref="ICellConverter{T}"/> of <c>T</c>.
    /// Returns null for a collection that the converter doesn't convert as a whole: it converts the
    /// collection's elements instead.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
#if NET5_0_OR_GREATER
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(NullableCellConverter<>))]
#endif
    private static ICellMapper? CreateConverterMapper<T>(object converter, MemberInfo member)
    {
        if (converter is ICellConverter<T> typedConverter)
        {
            return new CellConverterMapper<T>(typedConverter);
        }

        if (Nullable.GetUnderlyingType(typeof(T)) is { } underlyingType &&
            typeof(ICellConverter<>).MakeGenericType(underlyingType).IsInstanceOfType(converter))
        {
            var nullableConverter = (ICellConverter<T>)Activator.CreateInstance(typeof(NullableCellConverter<>).MakeGenericType(underlyingType), converter)!;
            return new CellConverterMapper<T>(nullableConverter);
        }

        if (typeof(T) != typeof(string) && typeof(System.Collections.IEnumerable).IsAssignableFrom(typeof(T)))
        {
            return null;
        }

        throw new ExcelMappingException($"The converter \"{converter.GetType()}\" given by ExcelConverterAttribute on member \"{member.Name}\" doesn't implement ICellConverter<{typeof(T)}>.");
    }

    internal static void ModifyMappers(IValuePipeline pipeline, MemberInfo member)
    {
        // If the member has a ExcelFormats attribute, modify the mappers.
        if (MemberAttributes.Get<ExcelFormatsAttribute>(member) is { } formatsAttribute)
        {
            foreach (var mapper in pipeline.Mappers.OfType<IFormatsCellMapper>())
            {
                mapper.Formats = [.. formatsAttribute.Formats];
            }
        }

        // If the member has a ExcelNumberStyle attribute, modify the mappers.
        if (MemberAttributes.Get<ExcelNumberStyleAttribute>(member) is { } numberStyleAttribute)
        {
            foreach (var mapper in pipeline.Mappers.OfType<INumberStyleCellMapper>())
            {
                mapper.Style = numberStyleAttribute.Style;
            }
        }

        // If the member has a ExcelUri attribute, modify the mappers.
        if (MemberAttributes.Get<ExcelUriAttribute>(member) is { } uriKindAttribute)
        {
            foreach (var mapper in pipeline.Mappers.OfType<UriMapper>())
            {
                mapper.UriKind = uriKindAttribute.UriKind;
            }
        }

        if (MemberAttributes.Get<ExcelOffsetAttribute>(member) is { } offsetAttribute)
        {
            foreach (var mapper in pipeline.Mappers.OfType<DateTimeOffsetMapper>())
            {
                mapper.Offset = offsetAttribute.Offset;
            }
        }
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    internal static IFallbackItem? GetDefaultEmptyValueFallback<T>(MemberInfo member)
    {
        // If the member has a ExcelDefaultValue attribute, add the fallback.
        if (MemberAttributes.Get<ExcelDefaultValueAttribute>(member) is { } defaultValueAttribute)
        {
            return new FixedValueFallback(defaultValueAttribute.Value);
        }

        // If the member has a ExcelEmptyFallback attribute, add the fallback.
        if (MemberAttributes.Get<ExcelEmptyFallbackAttribute>(member) is { } emptyFallback)
        {
            return CreateFallbackItem<T>(Activator.CreateInstance(emptyFallback.Type, emptyFallback.ConstructorArguments)!, member, nameof(ExcelEmptyFallbackAttribute));
        }

        return null;
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    internal static IFallbackItem? GetDefaultInvalidValueFallback<T>(MemberInfo member)
    {
        // If the member has a ExcelInvalidValue attribute, add the fallback.
        if (MemberAttributes.Get<ExcelInvalidValueAttribute>(member) is { } invalidValueAttribute)
        {
            return new FixedValueFallback(invalidValueAttribute.Value);
        }

        // If the member has a ExcelInvalidFallback attribute, add the fallback.
        if (MemberAttributes.Get<ExcelInvalidFallbackAttribute>(member) is { } invalidFallback)
        {
            return CreateFallbackItem<T>(Activator.CreateInstance(invalidFallback.Type, invalidFallback.ConstructorArguments)!, member, nameof(ExcelInvalidFallbackAttribute));
        }

        return null;
    }

    /// <summary>
    /// Creates the fallback for an object given by <see cref="ExcelEmptyFallbackAttribute"/> or
    /// <see cref="ExcelInvalidFallbackAttribute"/>: an <see cref="IFallbackItem"/>, an
    /// <see cref="ICellFallback{T}"/>, or for a <c>T?</c> member, an <see cref="ICellFallback{T}"/> of <c>T</c>.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
#if NET5_0_OR_GREATER
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(NullableCellFallback<>))]
#endif
    private static IFallbackItem CreateFallbackItem<T>(object fallback, MemberInfo member, string attributeName)
    {
        if (fallback is ICellFallback<T> typedFallback)
        {
            return new CellFallbackItem<T>(typedFallback);
        }

        if (Nullable.GetUnderlyingType(typeof(T)) is { } underlyingType &&
            typeof(ICellFallback<>).MakeGenericType(underlyingType).IsInstanceOfType(fallback))
        {
            var nullableFallback = (ICellFallback<T>)Activator.CreateInstance(typeof(NullableCellFallback<>).MakeGenericType(underlyingType), fallback)!;
            return new CellFallbackItem<T>(nullableFallback);
        }

        if (fallback is IFallbackItem fallbackItem)
        {
            return fallbackItem;
        }

        throw new ExcelMappingException($"The fallback \"{fallback.GetType()}\" given by {attributeName} on member \"{member.Name}\" doesn't implement ICellFallback<{typeof(T)}>.");
    }

    internal static void AddTransformers(IValuePipeline pipeline, MemberInfo member)
    {
        // If the member has a ExcelTrimString attribute, add the transformer.
        if (MemberAttributes.Get<ExcelTrimStringAttribute>(member) is { } trimStringAttribute)
        {
            pipeline.Transformers.Add(new TrimStringCellTransformer());
        }

        // If the member has ExcelTransformer attributes, add the transformers.
        if (MemberAttributes.GetMany<ExcelTransformerAttribute>(member) is { } transformAttributes)
        {
            foreach (var transformAttribute in transformAttributes)
            {
                var transformer = Activator.CreateInstance(transformAttribute.Type, transformAttribute.ConstructorArguments)!;
                pipeline.Transformers.Add(transformer is ICellTransformer cellTransformer
                    ? new CellTransformerAdapter(cellTransformer)
                    : (Abstractions.ICellTransformer)transformer);
            }
        }
    }

    internal static void ApplyMemberAttributes(IToOneMap map, MemberInfo member)
    {
        // If a member has an ExcelOptional attribute, mark it as optional.
        if (MemberAttributes.IsDefined<ExcelOptionalAttribute>(member))
        {
            map.Optional = true;
        }
        // If a member has an ExcelPreserveFormatting attribute, mark it as preserving formatting.
        if (MemberAttributes.IsDefined<ExcelPreserveFormattingAttribute>(member))
        {
            map.PreserveFormatting = true;
        }
    }
}
