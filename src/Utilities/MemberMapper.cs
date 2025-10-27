using System.Linq;
using System.Reflection;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;

namespace ExcelMapper.Utilities;

internal static class MemberMapper
{
    internal static ICellReaderFactory GetDefaultCellReaderFactory(MemberInfo member)
    {
        var columnNameAttributes = member.GetCustomAttributes<ExcelColumnNameAttribute>().ToArray();
        // A single [ExcelColumnName] attribute represents one column.
        if (columnNameAttributes.Length == 1)
        {
            return new ColumnNameReaderFactory(columnNameAttributes[0].Name, columnNameAttributes[0].Comparison);
        }
        // Multiple [ExcelColumnName] attributes still represents one column, but multiple options.
        else if (columnNameAttributes.Length > 1)
        {
            return new CompositeCellsReaderFactory(
                [.. columnNameAttributes.Select(c => new ColumnNameReaderFactory(c.Name, c.Comparison))]);
        }

        // [ExcelColumnNames] attributes still represents one column, but multiple options.
        var columnNamesAttribute = member.GetCustomAttribute<ExcelColumnNamesAttribute>();
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesReaderFactory(columnNamesAttribute.Names, columnNamesAttribute.Comparison);
        }

        // A single [ExcelColumnNameMatching] attributes still represents one column, but multiple options.
        var columnNameMatchingAttribute = member.GetCustomAttribute<ExcelColumnMatchingAttribute>();
        if (columnNameMatchingAttribute != null)
        {
            var matcher = (IExcelColumnMatcher)Activator.CreateInstance(columnNameMatchingAttribute.Type, columnNameMatchingAttribute.ConstructorArguments)!;
            return new ColumnsMatchingReaderFactory(matcher);
        }

        // A single [ExcelColumnIndex] attribute represents one column.
        var colummnIndexAttributes = member.GetCustomAttributes<ExcelColumnIndexAttribute>().ToArray();
        if (colummnIndexAttributes.Length == 1)
        {
            return new ColumnIndexReaderFactory(colummnIndexAttributes[0].Index);
        }
        // Multiple [ExcelColumnIndex] attributes still represents one column, but multiple options.
        else if (colummnIndexAttributes.Length > 1)
        {
            return new ColumnIndicesReaderFactory([.. colummnIndexAttributes.Select(c => c.Index)]);
        }

        // [ExcelColumnIndices] attributes still represents one column, but multiple options.
        var columnIndicesAttribute = member.GetCustomAttribute<ExcelColumnIndicesAttribute>();
        if (columnIndicesAttribute != null)
        {
            return new ColumnIndicesReaderFactory(columnIndicesAttribute.Indices);
        }

        return new ColumnNameReaderFactory(member.Name);
    }

    internal static ICellsReaderFactory GetDefaultSplitCellsReaderFactory(MemberInfo? member, ICellReaderFactory innerReaderFactory)
    {
        if (member != null)
        {
            // If the member has a ExcelSeparator attribute, use that separator.
            if (member.GetCustomAttribute<ExcelSeparatorsAttribute>() is { } separatorsAttribute)
            {
                if (separatorsAttribute.StringSeparators is { } stringSeparators)
                {
                    return new StringSplitReaderFactory(innerReaderFactory)
                    {
                        Separators = stringSeparators,
                        Options = separatorsAttribute.Options
                    };
                }

                return new CharSplitReaderFactory(innerReaderFactory)
                {
                    Separators = separatorsAttribute.CharSeparators!,
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
        var columnNamesAttribute = member.GetCustomAttribute<ExcelColumnNamesAttribute>();
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesReaderFactory(columnNamesAttribute.Names, columnNamesAttribute.Comparison);
        }

        // [ExcelColumnsMatchingAttribute] attributes represent multiple columns.
        var columnNameMatchingAttribute = member.GetCustomAttribute<ExcelColumnsMatchingAttribute>();
        if (columnNameMatchingAttribute != null)
        {
            var matcher = (IExcelColumnMatcher)Activator.CreateInstance(columnNameMatchingAttribute.Type, columnNameMatchingAttribute.ConstructorArguments)!;
            return new ColumnsMatchingReaderFactory(matcher);
        }

        // [ExcelColumnIndices] attributes represents multiple columns.
        var columnIndicesAttribute = member.GetCustomAttribute<ExcelColumnIndicesAttribute>();
        if (columnIndicesAttribute != null)
        {
            return new ColumnIndicesReaderFactory(columnIndicesAttribute.Indices);
        }

        return null;
    }

    internal static bool AddMappers(IValuePipeline pipeline, MemberInfo member)
    {
        var addDefaultMappers = true;

        // If the member has ExcelMapper attributes, add the mappers.
        var mapperAttributes = member.GetCustomAttributes<ExcelMapperAttribute>().ToArray();
        if (mapperAttributes.Length > 0)
        {
            foreach (var mapperAttribute in mapperAttributes)
            {
                var mapper = (ICellMapper)Activator.CreateInstance(mapperAttribute.Type, mapperAttribute.ConstructorArguments)!;
                pipeline.Mappers.Add(mapper);
            }

            // Since explicit mappers were added, do not add default mappers.
            addDefaultMappers = false;
        }

        // If the member has any ExcelMappingDictionary attributes, add a MappingDictionaryMapper.
        var mappingDictionaryAttributes = member.GetCustomAttributes<ExcelMappingDictionaryAttribute>().ToArray();
        if (mappingDictionaryAttributes.Length > 0)
        {
            // Build the mapping dictionary.
            var mappingDictionary = new Dictionary<string, object?>(mappingDictionaryAttributes.Length);
            foreach (var mappingDictionaryAttribute in mappingDictionaryAttributes)
            {
                mappingDictionary.Add(mappingDictionaryAttribute.Value, mappingDictionaryAttribute.MappedValue);
            }

            // If the member has a ExcelMappingDictionaryComparer attribute, get the comparer.
            IEqualityComparer<string>? comparer = null;
            if (member.GetCustomAttribute<ExcelMappingDictionaryComparerAttribute>() is { } comparerAttribute)
            {
                comparer = StringComparer.FromComparison(comparerAttribute.Comparison);
            }

            // If the member has a ExcelMappingDictionaryBehavior attribute, get the behavior.
            var behavior = MappingDictionaryMapperBehavior.Optional;
            if (member.GetCustomAttribute<ExcelMappingDictionaryBehaviorAttribute>() is { } behaviorAttribute)
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

    internal static void ModifyMappers(IValuePipeline pipeline, MemberInfo member)
    {
        // If the member has a ExcelFormats attribute, modify the mappers.
        if (member.GetCustomAttribute<ExcelFormatsAttribute>() is { } formatsAttribute)
        {
            foreach (var mapper in pipeline.Mappers.OfType<IFormatsCellMapper>())
            {
                mapper.Formats = formatsAttribute.Formats;
            }
        }
    }

    internal static IFallbackItem? GetDefaultEmptyValueFallback(MemberInfo member)
    {
        // If the member has a ExcelDefaultValue attribute, add the fallback.
        if (member.GetCustomAttribute<ExcelDefaultValueAttribute>() is { } defaultValueAttribute)
        {
            return new FixedValueFallback(defaultValueAttribute.Value);
        }

        // If the member has a ExcelEmptyFallback attribute, add the fallback.
        if (member.GetCustomAttribute<ExcelEmptyFallbackAttribute>() is { } emptyFallback)
        {
            return (IFallbackItem)Activator.CreateInstance(emptyFallback.Type, emptyFallback.ConstructorArguments)!;
        }

        return null;
    }

    internal static IFallbackItem? GetDefaultInvalidValueFallback(MemberInfo member)
    {
        // If the member has a ExcelInvalidValue attribute, add the fallback.
        if (member.GetCustomAttribute<ExcelInvalidValueAttribute>() is { } invalidValueAttribute)
        {
            return new FixedValueFallback(invalidValueAttribute.Value);
        }

        // If the member has a ExcelInvalidFallback attribute, add the fallback.
        if (member.GetCustomAttribute<ExcelInvalidFallbackAttribute>() is { } invalidFallback)
        {
            return (IFallbackItem)Activator.CreateInstance(invalidFallback.Type, invalidFallback.ConstructorArguments)!;
        }

        return null;
    }

    internal static void AddTransformers(IValuePipeline pipeline, MemberInfo member)
    {
        // If the member has a ExcelTrimString attribute, add the transformer.
        if (member.GetCustomAttribute<ExcelTrimStringAttribute>() is { } trimStringAttribute)
        {
            pipeline.Transformers.Add(new TrimStringCellTransformer());
        }

        // If the member has ExcelTransformer attributes, add the transformers.
        if (member.GetCustomAttributes<ExcelTransformerAttribute>() is { } transformAttributes)
        {
            foreach (var transformAttribute in transformAttributes)
            {
                var transformer = (ICellTransformer)Activator.CreateInstance(transformAttribute.Type, transformAttribute.ConstructorArguments)!;
                pipeline.Transformers.Add(transformer);
            }
        }
    }

    internal static void ApplyMemberAttributes(IToOneMap map, MemberInfo member)
    {
        // If a member has an ExcelOptional attribute, mark it as optional.
        if (Attribute.IsDefined(member, typeof(ExcelOptionalAttribute)))
        {
            map.Optional = true;
        }
        // If a member has an ExcelPreserveFormatting attribute, mark it as preserving formatting.
        if (Attribute.IsDefined(member, typeof(ExcelPreserveFormattingAttribute)))
        {
            map.PreserveFormatting = true;
        }
    }
}
