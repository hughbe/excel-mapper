using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;

namespace ExcelMapper.AutoMapping;

internal static class MemberMapper
{
    /// <summary>
    /// Gets the column a single value is read from, from the member's attributes, or the column with
    /// the member's name.
    /// </summary>
    internal static Columns GetDefaultColumns(MemberInfo member)
    {
        var columnNameAttributes = MemberAttributes.GetMany<ExcelColumnNameAttribute>(member);
        if (columnNameAttributes.Length == 1)
        {
            return new ColumnNameColumns(columnNameAttributes[0].Name, columnNameAttributes[0].Comparison);
        }

        if (columnNameAttributes.Length > 1)
        {
            var columns = new ColumnNameColumns[columnNameAttributes.Length];
            for (int i = 0; i < columns.Length; i++)
            {
                columns[i] = new ColumnNameColumns(columnNameAttributes[i].Name, columnNameAttributes[i].Comparison);
            }

            return new FirstColumnNameColumns(columns);
        }

        var columnNamesAttribute = MemberAttributes.Get<ExcelColumnNamesAttribute>(member);
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesColumns(columnNamesAttribute.Names, columnNamesAttribute.Comparison);
        }

        // A single [ExcelColumnMatching] attribute still reads one column, the first that matches.
        var columnMatchingAttribute = MemberAttributes.Get<ExcelColumnMatchingAttribute>(member);
        if (columnMatchingAttribute != null)
        {
            return CreateMatchingColumns(columnMatchingAttribute.Type, columnMatchingAttribute.ConstructorArguments, nameof(ExcelColumnMatchingAttribute), member);
        }

        var columnIndexAttributes = MemberAttributes.GetMany<ExcelColumnIndexAttribute>(member);
        if (columnIndexAttributes.Length == 1)
        {
            return new ColumnIndexColumns(columnIndexAttributes[0].Index);
        }

        if (columnIndexAttributes.Length > 1)
        {
            var indices = new int[columnIndexAttributes.Length];
            for (int i = 0; i < indices.Length; i++)
            {
                indices[i] = columnIndexAttributes[i].Index;
            }

            return new ColumnIndicesColumns(indices);
        }

        // [ExcelColumnIndices] still reads one column, the first the sheet has.
        var columnIndicesAttribute = MemberAttributes.Get<ExcelColumnIndicesAttribute>(member);
        if (columnIndicesAttribute != null)
        {
            return new ColumnIndicesColumns(columnIndicesAttribute.Indices);
        }

        return new ColumnNameColumns(member.Name);
    }

    /// <summary>
    /// Gets how a collection member is split from one cell: the cell is in <paramref name="cell"/>,
    /// with the separators of the member's <see cref="ExcelSeparatorsAttribute"/>, or commas.
    /// </summary>
    internal static SplitColumns GetDefaultSplitColumns(MemberInfo? member, Columns cell)
    {
        if (member != null && MemberAttributes.Get<ExcelSeparatorsAttribute>(member) is { } separatorsAttribute)
        {
            if (separatorsAttribute.StringSeparators is { } stringSeparators)
            {
                return new SplitColumns(cell, new StringSplitter
                {
                    Separators = [.. stringSeparators],
                    Options = separatorsAttribute.Options
                });
            }

            return new SplitColumns(cell, new CharSplitter
            {
                Separators = [.. separatorsAttribute.CharSeparators!],
                Options = separatorsAttribute.Options
            });
        }

        return new SplitColumns(cell, new CharSplitter());
    }

    /// <summary>
    /// Gets the columns a collection or dictionary member reads from its attributes, or null if it
    /// has none (it's then split from one cell, or for a dictionary, reads every column).
    /// </summary>
    internal static Columns? GetDefaultCellsColumns(MemberInfo? member)
    {
        // A whole row read as a collection or dictionary reads every column.
        if (member == null)
        {
            return AllColumns.Instance;
        }

        var columnNamesAttribute = MemberAttributes.Get<ExcelColumnNamesAttribute>(member);
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesColumns(columnNamesAttribute.Names, columnNamesAttribute.Comparison);
        }

        var columnsMatchingAttribute = MemberAttributes.Get<ExcelColumnsMatchingAttribute>(member);
        if (columnsMatchingAttribute != null)
        {
            return CreateMatchingColumns(columnsMatchingAttribute.Type, columnsMatchingAttribute.ConstructorArguments, nameof(ExcelColumnsMatchingAttribute), member);
        }

        var columnIndicesAttribute = MemberAttributes.Get<ExcelColumnIndicesAttribute>(member);
        if (columnIndicesAttribute != null)
        {
            return new ColumnIndicesColumns(columnIndicesAttribute.Indices);
        }

        return null;
    }

    // The type of [ExcelColumnMatching] and [ExcelColumnsMatching] is an IColumnSelector, or for a
    // pattern, RegexColumnMatcher.
    private static Columns CreateMatchingColumns([DynamicallyAccessedMembers(TrimmingConstants.ActivatedType)] Type type, object?[]? arguments, string attributeName, MemberInfo member)
        => MemberAttributes.CreateInstance(type, arguments, attributeName, member) switch
        {
            RegexColumnMatcher regexMatcher => new MatchingColumns(regexMatcher.Regex),
            var selector => new SelectorColumns((IColumnSelector)selector!),
        };

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    internal static bool AddMappers<T>(CellPipeline<T> pipeline, MemberInfo member)
    {
        var addDefaultMappers = true;

        bool hasMapperAttributes = false;

        // If the member has ExcelConverter attributes, add their converters.
        foreach (var converterAttribute in MemberAttributes.GetMany<ExcelConverterAttribute>(member))
        {
            var converter = MemberAttributes.CreateInstance(converterAttribute.Type, converterAttribute.ConstructorArguments, nameof(ExcelConverterAttribute), member);
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
        if (mappingDictionaryAttributes.Length != 0)
        {
            // If the member has a ExcelMappingDictionaryComparer attribute, get the comparer.
            IEqualityComparer<string>? comparer = null;
            if (MemberAttributes.Get<ExcelMappingDictionaryComparerAttribute>(member) is { } comparerAttribute)
            {
                comparer = StringComparerUtilities.FromComparison(comparerAttribute.Comparison);
            }

            // Keyed with the comparer, so that text the comparer treats as the same is a duplicate.
            var mappingDictionary = new Dictionary<string, T>(comparer);
            foreach (var mappingDictionaryAttribute in mappingDictionaryAttributes)
            {
                var text = mappingDictionaryAttribute.Value;
                if (mappingDictionary.ContainsKey(text))
                {
                    throw new ExcelMappingException($"Member \"{member.Name}\" has more than one ExcelMappingDictionaryAttribute for \"{text}\".");
                }

                mappingDictionary.Add(text, MemberAttributes.ConvertValue<T>(mappingDictionaryAttribute.MappedValue, nameof(ExcelMappingDictionaryAttribute), member, text));
            }

            // If the member has a ExcelMappingDictionaryBehavior attribute, get the behavior.
            var behavior = MappingDictionaryBehavior.Optional;
            if (MemberAttributes.Get<ExcelMappingDictionaryBehaviorAttribute>(member) is { } behaviorAttribute)
            {
                behavior = behaviorAttribute.Behavior;
            }

            pipeline.Mappers.Add(new MappingDictionaryMapper<T>(mappingDictionary, comparer, behavior));

            // If the dictionary mapper was added as required, do not add default mappers.
            if (behavior == MappingDictionaryBehavior.Required)
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
    private static ICellMapper<T>? CreateConverterMapper<T>(object converter, MemberInfo member)
    {
        if ((converter as ICellConverter<T> ?? TypeRule<T>.Instance.NullableAdapter?.AdaptConverter(converter)) is { } typedConverter)
        {
            return new CellConverterMapper<T>(typedConverter);
        }

        if (typeof(T) != typeof(string) && typeof(System.Collections.IEnumerable).IsAssignableFrom(typeof(T)))
        {
            return null;
        }

        throw new ExcelMappingException($"The converter {TypeNames.Get(converter.GetType())} given by ExcelConverterAttribute on member \"{member.Name}\" doesn't implement ICellConverter<{TypeNames.Get(typeof(T))}>.");
    }

    /// <summary>
    /// Applies the member's attributes that change how its built-in conversion parses text
    /// (<see cref="ExcelFormatsAttribute"/>, <see cref="ExcelNumberStyleAttribute"/>,
    /// <see cref="ExcelUriKindAttribute"/> and <see cref="ExcelOffsetAttribute"/>) to the pipeline's
    /// options.
    /// </summary>
    internal static void ApplyConversionOptions(CellPipeline pipeline, MemberInfo member)
    {
        var options = pipeline.Options;
        if (MemberAttributes.Get<ExcelFormatsAttribute>(member) is { } formatsAttribute)
        {
            options.Formats = [.. formatsAttribute.Formats];
        }

        if (MemberAttributes.Get<ExcelNumberStyleAttribute>(member) is { } numberStyleAttribute)
        {
            options.NumberStyle = numberStyleAttribute.Style;
        }

        if (MemberAttributes.Get<ExcelUriKindAttribute>(member) is { } uriKindAttribute)
        {
            options.UriKind = uriKindAttribute.UriKind;
        }

        if (MemberAttributes.Get<ExcelOffsetAttribute>(member) is { } offsetAttribute)
        {
            options.Offset = offsetAttribute.Offset;
        }
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    internal static ICellFallback<T>? GetDefaultEmptyValueFallback<T>(MemberInfo member)
    {
        // If the member has a ExcelDefaultValue attribute, add the fallback.
        if (MemberAttributes.Get<ExcelDefaultValueAttribute>(member) is { } defaultValueAttribute)
        {
            return new FixedValueFallback<T>(MemberAttributes.ConvertValue<T>(defaultValueAttribute.Value, nameof(ExcelDefaultValueAttribute), member));
        }

        // If the member has a ExcelEmptyFallback attribute, add the fallback.
        if (MemberAttributes.Get<ExcelEmptyFallbackAttribute>(member) is { } emptyFallback)
        {
            return CreateFallback<T>(MemberAttributes.CreateInstance(emptyFallback.Type, emptyFallback.ConstructorArguments, nameof(ExcelEmptyFallbackAttribute), member), member, nameof(ExcelEmptyFallbackAttribute));
        }

        return null;
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    internal static ICellFallback<T>? GetDefaultInvalidValueFallback<T>(MemberInfo member)
    {
        // If the member has a ExcelInvalidValue attribute, add the fallback.
        if (MemberAttributes.Get<ExcelInvalidValueAttribute>(member) is { } invalidValueAttribute)
        {
            return new FixedValueFallback<T>(MemberAttributes.ConvertValue<T>(invalidValueAttribute.Value, nameof(ExcelInvalidValueAttribute), member));
        }

        // If the member has a ExcelInvalidFallback attribute, add the fallback.
        if (MemberAttributes.Get<ExcelInvalidFallbackAttribute>(member) is { } invalidFallback)
        {
            return CreateFallback<T>(MemberAttributes.CreateInstance(invalidFallback.Type, invalidFallback.ConstructorArguments, nameof(ExcelInvalidFallbackAttribute), member), member, nameof(ExcelInvalidFallbackAttribute));
        }

        return null;
    }

    /// <summary>
    /// Creates the fallback for an object given by <see cref="ExcelEmptyFallbackAttribute"/> or
    /// <see cref="ExcelInvalidFallbackAttribute"/>: an <see cref="ICellFallback{T}"/>, or for a <c>T?</c> member, an <see cref="ICellFallback{T}"/> of <c>T</c>.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    private static ICellFallback<T> CreateFallback<T>(object fallback, MemberInfo member, string attributeName)
    {
        if ((fallback as ICellFallback<T> ?? TypeRule<T>.Instance.NullableAdapter?.AdaptFallback(fallback)) is { } typedFallback)
        {
            return typedFallback;
        }

        throw new ExcelMappingException($"The fallback {TypeNames.Get(fallback.GetType())} given by {attributeName} on member \"{member.Name}\" doesn't implement ICellFallback<{TypeNames.Get(typeof(T))}>.");
    }

    internal static void AddTransformers(CellPipeline pipeline, MemberInfo member)
    {
        // If the member has a ExcelTrimString attribute, add the transformer.
        if (MemberAttributes.Get<ExcelTrimStringAttribute>(member) is { } trimStringAttribute)
        {
            pipeline.Transformers.Add(TrimStringCellTransformer.Instance);
        }

        // If the member has ExcelTransformer attributes, add the transformers.
        if (MemberAttributes.GetMany<ExcelTransformerAttribute>(member) is { } transformAttributes)
        {
            foreach (var transformAttribute in transformAttributes)
            {
                var transformer = MemberAttributes.CreateInstance(transformAttribute.Type, transformAttribute.ConstructorArguments, nameof(ExcelTransformerAttribute), member);
                pipeline.Transformers.Add((ICellTransformer)transformer);
            }
        }
    }

    internal static void ApplyMemberAttributes(IColumnsMap map, MemberInfo member)
    {
        // If a member has an ExcelOptional attribute, mark it as optional.
        if (MemberAttributes.IsDefined<ExcelOptionalAttribute>(member))
        {
            map.Settings.Optional = true;
        }
        // If a member has an ExcelPreserveFormatting attribute, mark it as preserving formatting.
        if (MemberAttributes.IsDefined<ExcelPreserveFormattingAttribute>(member))
        {
            map.Settings.PreserveFormatting = true;
        }
    }
}
