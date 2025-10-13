using System;
using System.Collections;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Factories;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

public static class AutoMapper
{
    private static MethodInfo? s_tryAutoMapMemberMethod;
    private static MethodInfo TryAutoMapMemberMethod => s_tryAutoMapMemberMethod ??= typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryAutoMapMember))!;

    /// <summary>
    /// Creates a map to assign a value to a specific member.
    /// </summary>
    /// <typeparam name="TMember">The target type.</typeparam>
    /// <param name="member"></param>
    /// <param name="emptyValueStrategy">The behaviour if the value is empty.</param>
    /// <param name="map">The pipeline.</param>
    /// <returns>True if the member is able to be mapped.</returns>
    private static bool TryAutoMapMember<TMember>(MemberInfo member, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        // First, check if this is a well-known type (e.g. string/int).
        // This is a simple conversion from the cell's value to the type.
        if (TryCreatePrimitiveMap(member, emptyValueStrategy, true, out OneToOneMap<TMember>? singleMap))
        {
            map = singleMap;
            return true;
        }

        // Secondly, check if this is a dictionary.
        // This requires converting each value to the value type of the collection.
        if (TryCreateDictionaryMap<TMember>(member, emptyValueStrategy, isAutoMapping: true, out var dictionaryMap))
        {
            map = dictionaryMap;
            return true;
        }

        // Thirdly, check if this is a collection (e.g. array, list).
        // This requires converting each value to the element type of the collection.
        if (TryCreateSplitMap(member, emptyValueStrategy, out var multiMap))
        {
            map = multiMap;
            return true;
        }

        // Fourthly, check if this is an object.
        // This requires converting each member and setting it on the object.
        if (TryCreateObjectMap<TMember>(emptyValueStrategy, out var objectMap))
        {
            map = objectMap;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryCreatePrimitivePipeline<T>(FallbackStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out ValuePipeline<T>? pipeline)
    {
        if (!TryGetWellKnownMap<T>(emptyValueStrategy, out var mapper, out var emptyFallback, out var invalidFallback))
        {
            if (isAutoMapping)
            {
                pipeline = null;
                return false;
            }
        }

        pipeline = new ValuePipeline<T>();
        if (mapper != null)
        {
            pipeline.AddCellValueMapper(mapper);
        }
        if (emptyFallback != null)
        {
            pipeline.EmptyFallback = emptyFallback;
        }
        if (invalidFallback != null)
        {
            pipeline.InvalidFallback = invalidFallback;
        }
        return true;
    }

    private static bool TryCreatePrimitiveMap<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out OneToOneMap<T>? map)
    {
        map = CreateMemberMap<T>(member, emptyValueStrategy, isAutoMapping);
        return map != null;
    }

    internal static OneToOneMap<T>? CreateMemberMap<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, bool isAutoMapping)
    {
        if (!TryGetWellKnownMap<T>(emptyValueStrategy, out ICellMapper? mapper, out IFallbackItem? emptyFallback, out IFallbackItem? invalidFallback))
        {
            // Cannot auto map an unsupported primitive.
            // But allow `Map(o => o.Value)` as the user can define a custom converter after
            // creating the map.
            // If the user doesn't add any mappers to the property, an ExcelMappingException
            // will be thrown at the point of registering the class map.
            if (isAutoMapping)
            {
                return null;
            }
        }

        var defaultValueAttribute = member.GetCustomAttribute<ExcelDefaultValueAttribute>();
        if (defaultValueAttribute != null)
        {
            emptyFallback = new FixedValueFallback(defaultValueAttribute.Value);
        }

        var defaultReader = GetDefaultCellReaderFactory(member);
        var map = new OneToOneMap<T>(defaultReader);

        if (mapper != null)
        {
            map.AddCellValueMapper(mapper);
        }
        if (emptyFallback != null)
        {
            map.EmptyFallback = emptyFallback;
        }
        if (invalidFallback != null)
        {
            map.InvalidFallback = invalidFallback;
        }
        ApplyMemberAttributesToMap(member, map);

        return map;
    }

    private static ICellReaderFactory GetDefaultCellReaderFactory(MemberInfo member)
    {
        var columnNameAttributes = member.GetCustomAttributes<ExcelColumnNameAttribute>().ToArray();
        // A single [ExcelColumnName] attribute represents one column.
        if (columnNameAttributes.Length == 1)
        {
            return new ColumnNameReaderFactory(columnNameAttributes[0].Name);
        }
        // Multiple [ExcelColumnName] attributes still represents one column, but multiple options.
        else if (columnNameAttributes.Length > 1)
        {
            return new ColumnNamesReaderFactory([.. columnNameAttributes.Select(c => c.Name)]);
        }

        // [ExcelColumnNames] attributes still represents one column, but multiple options.
        var columnNamesAttribute = member.GetCustomAttribute<ExcelColumnNamesAttribute>();
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesReaderFactory(columnNamesAttribute.Names);
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

    private static bool TryGetWellKnownMap<T>(
        FallbackStrategy emptyValueStrategy,
        [NotNullWhen(true)] out ICellMapper? mapper,
        [NotNullWhen(true)] out IFallbackItem? emptyFallback,
        [NotNullWhen(true)] out IFallbackItem? invalidFallback)
    {
        Type type = typeof(T).GetNullableTypeOrThis(out bool isNullable);
        Type[] interfaces = [.. type.GetTypeInfo().ImplementedInterfaces];

        IFallbackItem ReconcileFallback(FallbackStrategy strategyToPursue, bool isEmpty)
        {
            // Empty nullable values should be set to null.
            if (isEmpty && isNullable)
            {
                return new FixedValueFallback(null);
            }
            else if (strategyToPursue == FallbackStrategy.SetToDefaultValue || emptyValueStrategy == FallbackStrategy.SetToDefaultValue)
            {
                return new FixedValueFallback(type.DefaultValue());
            }
            else
            {
                Debug.Assert(emptyValueStrategy == FallbackStrategy.ThrowIfPrimitive);

                // The user specified that we should set to the default value if it was empty.
                return new ThrowFallback();
            }
        }

        // Set the default mapper for each well-known type.
        if (type == typeof(DateTime))
        {
            mapper = new DateTimeMapper();
            emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
            invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
        }
        else if (type == typeof(Guid))
        {
            mapper = new GuidMapper();
            emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
            invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
        }
        else if (type == typeof(bool))
        {
            mapper = new BoolMapper();
            emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
            invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
        }
        else if (type.GetTypeInfo().IsEnum)
        {
            mapper = new EnumMapper(type);
            emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
            invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
        }
        else if (type == typeof(string) || type == typeof(object) || type == typeof(IConvertible))
        {
            mapper = new StringMapper();
            emptyFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, isEmpty: true);
            invalidFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, isEmpty: false);
        }
        else if (type == typeof(Uri))
        {
            mapper = new UriMapper();
            emptyFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, isEmpty: true);
            invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
        }
        else if (interfaces.Any(t => t == typeof(IConvertible)))
        {
            mapper = new ChangeTypeMapper(type);
            emptyFallback = ReconcileFallback(isNullable ? FallbackStrategy.SetToDefaultValue : FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
            invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
        }
        else
        {
            mapper = null;
            emptyFallback = null;
            invalidFallback = null;
            return false;
        }

        return true;
    }
    
    private static MethodInfo? s_tryCreateSplitGenericMapMethod;
    private static MethodInfo TryCreateSplitGenericMapMethod => s_tryCreateSplitGenericMapMethod ??= typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryCreateSplitMapGeneric))!;

    internal static bool TryCreateSplitMap(MemberInfo member, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        var listType = member.MemberType();
        var elementType = listType.GetElementTypeOrEnumerableType();
        if (elementType == null)
        {
            map = null;
            return false;
        }

        var method = TryCreateSplitGenericMapMethod.MakeGenericMethod([elementType]);
        var parameters = new object?[] { member, listType, emptyValueStrategy, null };
        var result = (bool)method.InvokeUnwrapped(null, parameters)!;
        if (result)
        {
            map = (IMap)parameters[^1]!;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryCreateSplitMap<TElement>(MemberInfo member, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        var method = TryCreateSplitGenericMapMethod.MakeGenericMethod([typeof(TElement)]);
        var parameters = new object?[] { member, member.MemberType(), emptyValueStrategy, null };
        var result = (bool)method.InvokeUnwrapped(null, parameters)!;
        if (result)
        {
            map = (IMap)parameters[^1]!;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryCreateEnumerableMap<T>(FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        Type listType = typeof(T);
        var elementType = listType.GetElementTypeOrEnumerableType();
        if (elementType == null)
        {
            map = null;
            return false;
        }

        var method = TryCreateSplitGenericMapMethod.MakeGenericMethod([elementType]);
        var parameters = new object?[] { null, listType, emptyValueStrategy, null };
        var result = (bool)method.InvokeUnwrapped(null, parameters)!;
        if (result)
        {
            map = (IMap)parameters[^1]!;
            return true;
        }

        map = null;
        return false;
    }

    private static bool TryCreateSplitMapGeneric<TElement>(MemberInfo? member, Type listType, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ManyToOneEnumerableMap<TElement>? map)
    {
        // First, get the pipeline for the element. This is used to convert individual values
        // to be added to/included in the collection.
        if (!TryCreatePrimitivePipeline<TElement>(emptyValueStrategy, isAutoMapping: true, out var elementMapping))
        {
            map = null;
            return false;
        }

        // Secondly, find the right way of adding the converted value to the collection.
        if (!TryGetCreateEnumerableFactory<TElement>(listType, out var factory))
        {
            map = null;
            return false;
        }

        // Otherwise, fallback to splitting a single cell with the default comma separator.
        var defaultReaderFactory = GetDefaultCellsReaderFactory(member) ??  new CharSplitReaderFactory(GetDefaultCellReaderFactory(member!));
        map = new ManyToOneEnumerableMap<TElement>(defaultReaderFactory, elementMapping, factory);
        ApplyMemberAttributesToMap(member, map);

        return true;
    }

    private static ICellsReaderFactory? GetDefaultCellsReaderFactory(MemberInfo? member)
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
            return new ColumnNamesReaderFactory(columnNamesAttribute.Names);
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

    private static bool TryGetCreateEnumerableFactory<TElement>(Type listType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        if (listType == typeof(Array) || (listType.IsArray && listType.GetArrayRank() == 1))
        {
            result = new ArrayEnumerableFactory<TElement>();
            return true;
        }
        else if (listType == typeof(ImmutableArray<TElement>))
        {
            result = new ImmutableArrayEnumerableFactory<TElement>();
            return true;
        }
        else if (listType == typeof(ImmutableList<TElement>) || listType == typeof(IImmutableList<TElement>))
        {
            result = new ImmutableListEnumerableFactory<TElement>();
            return true;
        }
        else if (listType == typeof(ImmutableStack<TElement>) || listType == typeof(IImmutableStack<TElement>))
        {
            result = new ImmutableStackEnumerableFactory<TElement>();
            return true;
        }
        else if (listType == typeof(ImmutableQueue<TElement>) || listType == typeof(IImmutableQueue<TElement>))
        {
            result = new ImmutableQueueEnumerableFactory<TElement>();
            return true;
        }
        else if (listType == typeof(ImmutableSortedSet<TElement>))
        {
            result = new ImmutableSortedSetEnumerableFactory<TElement>();
            return true;
        }
        else if (listType == typeof(ImmutableHashSet<TElement>) || listType == typeof(IImmutableSet<TElement>))
        {
            result = new ImmutableHashSetEnumerableFactory<TElement>();
            return true;
        }
        else if (listType == typeof(FrozenSet<TElement>))
        {
            result = new FrozenSetEnumerableFactory<TElement>();
            return true;
        }
        else if (listType == typeof(ReadOnlyObservableCollection<TElement>))
        {
            result = new ReadOnlyObservableCollectionEnumerableFactory<TElement>();
            return true;
        }
        else if (listType.IsInterface)
        {
            // Add values by creating a list and assigning to the property.
            if (listType.IsAssignableFrom(typeof(List<TElement>).GetTypeInfo()))
            {
                result = new ListEnumerableFactory<TElement>();
                return true;
            }
            if (listType.IsAssignableFrom(typeof(HashSet<TElement>).GetTypeInfo()))
            {
                result = new HashSetEnumerableFactory<TElement>();
                return true;
            }
        }
        // Otheriwse, we have to create the type.
        else if (!listType.IsAbstract)
        {
            var hasDefaultConstructor = listType.GetConstructor([]) is not null;
            if (hasDefaultConstructor)
            {
                // Add values with through IList<TElement>.Add(TElement item).
                if (listType.ImplementsInterface(typeof(IList<TElement>)))
                {
                    result = new IListTImplementingEnumerableFactory<TElement>(listType);
                    return true;
                }

                // Add values with through ISet<TElement>.Add(TElement item).
                if (listType.ImplementsInterface(typeof(ISet<TElement>)))
                {
                    result = new ISetTImplementingEnumerableFactory<TElement>(listType);
                    return true;
                }

                // Add values with through ICollection<TElement>.Add(TElement item).
                if (listType.ImplementsInterface(typeof(ICollection<TElement>)))
                {
                    result = new ICollectionTImplementingEnumerableFactory<TElement>(listType);
                    return true;
                }

                // Add values with through IList.Add(TElement item).
                if (listType.ImplementsInterface(typeof(IList)))
                {
                    result = new IListImplementingEnumerableFactory<TElement>(listType);
                    return true;
                }
            }

            // Check if the type has .ctor(ISet<T>) such as ReadOnlySet.
            var ctor = listType.GetConstructor([typeof(ISet<TElement>)]);
            if (ctor != null)
            {
                result = new ConstructorSetEnumerableFactory<TElement>(listType);
                return true;
            }

            // Check if the type has .ctor(IList<T>) .ctor(ICollection) or .ctor(IEnumerable<T>) such as ReadOnlyCollection, Queue or Stack.
            ctor = listType.GetConstructor([typeof(IList<TElement>)])
                ?? listType.GetConstructor([typeof(IEnumerable<TElement>)])
                ?? listType.GetConstructor([typeof(ICollection)]);
            if (ctor != null)
            {
                result = new ConstructorEnumerableFactory<TElement>(listType);
                return true;
            }

            // Check if the type has Add(T) such as BlockingCollection.
            if (hasDefaultConstructor)
            {
                var addMethod = listType.GetMethod("Add", [typeof(TElement)]);
                if (addMethod != null)
                {
                    result = new AddEnumerableFactory<TElement>(listType);
                    return true;
                }
            }
        }

        result = default;
        return false;
    }

    private static MethodInfo? s_getOrCreateArrayIndexerMapGenericMethod;
    private static MethodInfo GetOrCreateArrayIndexerMapGenericMethod => s_getOrCreateArrayIndexerMapGenericMethod ??= typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(GetOrCreateArrayIndexerMapGeneric))!;

    internal static IMap GetOrCreateArrayIndexerMap(IMap parentMap, MemberInfo? member, Type arrayType, object? index, Type elementType)
    {
        var method = GetOrCreateArrayIndexerMapGenericMethod.MakeGenericMethod([elementType]);
        var parameters = new object?[] { parentMap, member, arrayType, index };
        return (IMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static ManyToOneEnumerableIndexerMapT<TElement> GetOrCreateArrayIndexerMapGeneric<TElement>(IMap parentMap, MemberInfo? member, Type arrayType, object? index)
    {
        if (GetExistingMap(parentMap, member, index) is { } existingMap)
        {
            if (existingMap is not ManyToOneEnumerableIndexerMapT<TElement> arrayIndexerMap)
            {
                throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
            }

            return arrayIndexerMap;
        }

        // Create a new array indexer map.
        if (!TryCreateArrayIndexerMapGeneric<TElement>(member, arrayType, GetEmptyValueStrategy(parentMap), out var mapObj))
        {
            throw new ExcelMappingException($"Could not map array of type \"{arrayType}\".");
        }

        AddExistingMap(parentMap, member, index, mapObj);
        return mapObj;
    }

    private static bool TryCreateArrayIndexerMapGeneric<TElement>(MemberInfo? member, Type listType, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ManyToOneEnumerableIndexerMapT<TElement>? map)
    {
        if (!TryGetCreateEnumerableFactory<TElement>(listType, out var factory))
        {
            map = null;
            return false;
        }

        // Get the column names/indices from the attributes on the member.
        map = new ManyToOneEnumerableIndexerMapT<TElement>(factory);
        return true;
    }
    private static MethodInfo? s_getOrCreateMultidimensionalIndexerMapGenericMethod;
    private static MethodInfo GetOrCreateMultidimensionalIndexerMapGenericMethod => s_getOrCreateMultidimensionalIndexerMapGenericMethod ??= typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(GetOrCreateMultidimensionalIndexerMapGeneric))!;

    internal static IMap GetOrCreateMultidimensionalIndexerMap(IMap parentMap, MemberInfo? member, Type arrayType, object? index, Type elementType)
    {
        var method = GetOrCreateMultidimensionalIndexerMapGenericMethod.MakeGenericMethod([elementType]);
        var parameters = new object?[] { parentMap, member, arrayType, index };
        return (IMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static ManyToOneMultidimensionalIndexerMapT<TElement> GetOrCreateMultidimensionalIndexerMapGeneric<TElement>(IMap parentMap, MemberInfo? member, Type arrayType, object? index)
    {
        if (GetExistingMap(parentMap, member, index) is { } existingMap)
        {
            if (existingMap is not ManyToOneMultidimensionalIndexerMapT<TElement> multidimensionalMap)
            {
                throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
            }

            return multidimensionalMap;
        }

        // Create a new array indexer map.
        var factory = new MultidimensionalArrayFactory<TElement>();
        var map = new ManyToOneMultidimensionalIndexerMapT<TElement>(factory);
        AddExistingMap(parentMap, member, index, map);
        return map;
    }

    private static MethodInfo? s_getOrCreateDictionaryIndexerMapGenericMethod;
    private static MethodInfo GetOrCreateDictionaryIndexerMapGenericMethod => s_getOrCreateDictionaryIndexerMapGenericMethod ??= typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(GetOrCreateDictionaryIndexerMapGeneric))!;

    internal static IDictionaryIndexerMap GetOrCreateDictionaryIndexerMap(IMap parentMap, MemberInfo? member, Type dictionaryType, object? index, Type keyType, Type valueType)
    {
        var method = GetOrCreateDictionaryIndexerMapGenericMethod.MakeGenericMethod([keyType, valueType]);
        var parameters = new object?[] { parentMap, member, dictionaryType, index };
        return (IDictionaryIndexerMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static ManyToOneDictionaryIndexerMapT<TKey, TValue> GetOrCreateDictionaryIndexerMapGeneric<TKey, TValue>(IMap parentMap, MemberInfo? member, Type dictionaryType, object? index) where TKey : notnull
    {
        if (GetExistingMap(parentMap, member, index) is { } existingMap)
        {
            if (existingMap is not ManyToOneDictionaryIndexerMapT<TKey, TValue> dictionaryIndexerMap)
            {
                throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
            }

            return dictionaryIndexerMap;
        }

        // Create a new dictionary indexer map.
        if (!TryCreateDictionaryIndexerMapGeneric<TKey, TValue>(member, dictionaryType, GetEmptyValueStrategy(parentMap), out var mapObj))
        {
            throw new ExcelMappingException($"Could not map dictionary of type \"{dictionaryType}\".");
        }
        
        AddExistingMap(parentMap, member, index, mapObj);
        return mapObj;
    }

    private static bool TryCreateDictionaryIndexerMapGeneric<TKey, TValue>(MemberInfo? member, Type dictionaryType, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ManyToOneDictionaryIndexerMapT<TKey, TValue>? map) where TKey : notnull
    {
        if (!TryGetCreateDictionaryFactory<TKey, TValue>(dictionaryType, out var factory))
        {
            map = null;
            return false;
        }

        map = new ManyToOneDictionaryIndexerMapT<TKey, TValue>(factory);
        return true;
    }

    internal static IMap CreateArrayIndexerElementMap(int index, Type valueType, FallbackStrategy emptyValueStrategy)
        => CreateIndexerElementMap(new ColumnIndexReaderFactory(index), valueType, emptyValueStrategy);

    private static ICellReaderFactory CreateDefaultDictionaryKeyReaderFactory(object key)
        => key switch
        {
            string keyString when keyString.Length == 0 => new ColumnIndexReaderFactory(0),
            string keyString => new ColumnNameReaderFactory(keyString),
            int keyIndex => new ColumnIndexReaderFactory(keyIndex),
            _ => new ColumnNameReaderFactory(key.ToString()!)
        };

    internal static IMap CreateDictionaryIndexerElementMap(object key, Type valueType, FallbackStrategy emptyValueStrategy)
        => CreateIndexerElementMap(CreateDefaultDictionaryKeyReaderFactory(key), valueType, emptyValueStrategy);

    private static MethodInfo? s_tryCreateIndexerElementMapGenericMethod;
    private static MethodInfo CreateIndexerElementMapGenericMethod => s_tryCreateIndexerElementMapGenericMethod ??= typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(CreateIndexerElementMapGeneric))!;

    private static IMap CreateIndexerElementMap(ICellReaderFactory defaultReaderFactory, Type valueType, FallbackStrategy emptyValueStrategy)
    {
        var method = CreateIndexerElementMapGenericMethod.MakeGenericMethod(valueType);
        var parameters = new object?[] { defaultReaderFactory, emptyValueStrategy };
        return (IMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static IMap CreateIndexerElementMapGeneric<T>(ICellReaderFactory defaultReaderFactory, FallbackStrategy emptyValueStrategy)
    {
        // Try to create a primitive map for the value type.
        _ = TryGetWellKnownMap<T>(emptyValueStrategy, out var mapper, out var emptyFallback, out var invalidFallback);

        var map = new OneToOneMap<T>(defaultReaderFactory);
        if (mapper != null)
        {
            map.AddCellValueMapper(mapper);
        }
        if (emptyFallback != null)
        {
            map.EmptyFallback = emptyFallback;
        }
        if (invalidFallback != null)
        {
            map.InvalidFallback = invalidFallback;
        }
        
        return map;
    }

    private static MethodInfo? s_tryCreateDictionaryMapGeneric;
    private static MethodInfo TryCreateDictionaryMapGeneric => s_tryCreateDictionaryMapGeneric ??= typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryCreateGenericDictionaryMap))!;

    private static bool TryCreateDictionaryMap<TMember>(MemberInfo? member, FallbackStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out IMap? map)
    {
        if (!TryGetDictionaryKeyValueType(typeof(TMember), out var keyType, out var valueType))
        {
            map = null;
            return false;
        }

        var method = TryCreateDictionaryMapGeneric.MakeGenericMethod(keyType, valueType);
        var parameters = new object?[] { member, typeof(TMember), emptyValueStrategy, isAutoMapping, null };
        var result = (bool)method.InvokeUnwrapped(null, parameters)!;
        if (result)
        {
            map = (IMap)parameters[^1]!;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryGetDictionaryKeyValueType(Type dictionaryType, [NotNullWhen(true)] out Type? keyType, [NotNullWhen(true)] out Type? valueType)
    {
        // We should be able to parse anything that implements IEnumerable<KeyValuePair<TKey, TValue>>
        if (dictionaryType.ImplementsGenericInterface(typeof(IEnumerable<>), out Type? keyValuePairType))
        {
            if (keyValuePairType.IsGenericType && keyValuePairType.GetGenericTypeDefinition() == typeof(KeyValuePair<,>))
            {
                Type[] arguments = keyValuePairType.GenericTypeArguments;
                keyType = arguments[0];
                valueType = arguments[1];
                return true;
            }
        }

        // Parse regular IDictionary before checking for Add methods.
        // IDictionary also has Add(object, object) but we want to use string keys by convention.
        if (dictionaryType == typeof(IDictionary) || dictionaryType.ImplementsInterface(typeof(IDictionary)))
        {
            keyType = typeof(string);
            valueType = typeof(object);
            return true;
        }

        // Check for types like StringDictionary that have Add(TKey, TValue) method and implement IEnumerable.
        // This handles dictionary-like types that don't implement IDictionary<TKey, TValue> or IDictionary.
        if (dictionaryType.ImplementsInterface(typeof(IEnumerable)))
        {
            var addMethods = dictionaryType.GetMethods()
                .Where(m => m.Name == "Add" && m.GetParameters().Length == 2)
                .ToArray();
            
            if (addMethods.Length == 1)
            {
                var parameters = addMethods[0].GetParameters();
                keyType = parameters[0].ParameterType;
                valueType = parameters[1].ParameterType;
                return true;
            }
        }
        
        keyType = null;
        valueType = null;
        return false;
    }

    internal static bool TryCreateGenericDictionaryMap<TKey, TValue>(MemberInfo? member, Type dictionaryType, FallbackStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out ManyToOneDictionaryMap<TKey, TValue>? map) where TKey : notnull
    {
        if (!TryCreatePrimitivePipeline<TValue>(emptyValueStrategy, isAutoMapping, out var valuePipeline))
        {
            map = null;
            return false;
        }

        if (!TryGetCreateDictionaryFactory<TKey, TValue>(dictionaryType, out var factory))
        {
            map = null;
            return false;
        }

        // Default to all columns.
        var defaultReaderFactory = GetDefaultCellsReaderFactory(member) ?? new AllColumnNamesReaderFactory();
        map = new ManyToOneDictionaryMap<TKey, TValue>(defaultReaderFactory, valuePipeline, factory);
        ApplyMemberAttributesToMap(member, map);
        return true;
    }

    private static bool TryGetCreateDictionaryFactory<TKey, TValue>(Type dictionaryType, [NotNullWhen(true)] out IDictionaryFactory<TKey, TValue>? result) where TKey : notnull
    {
        if (dictionaryType == typeof(ImmutableDictionary<TKey, TValue>) || dictionaryType == typeof(IImmutableDictionary<TKey, TValue>))
        {
            result = new ImmutableDictionaryFactory<TKey, TValue>();
            return true;
        }
        else if (dictionaryType == typeof(ImmutableSortedDictionary<TKey, TValue>))
        {
            result = new ImmutableSortedDictionaryFactory<TKey, TValue>();
            return true;
        }
        else if (dictionaryType == typeof(FrozenDictionary<TKey, TValue>))
        {
            result = new FrozenDictionaryFactory<TKey, TValue>();
            return true;
        }
        else if (dictionaryType.GetTypeInfo().IsInterface)
        {
            if (dictionaryType.GetTypeInfo().IsAssignableFrom(typeof(Dictionary<TKey, TValue>).GetTypeInfo()))
            {
                result = new DictionaryFactory<TKey, TValue>();
                return true;
            }
        }
        else if (dictionaryType.GetConstructor([]) is not null)
        {
            if (dictionaryType.ImplementsInterface(typeof(IDictionary<TKey, TValue>)))
            {
                result = new IDictionaryTImplementingFactory<TKey, TValue>(dictionaryType);
                return true;
            }
            else if (dictionaryType.ImplementsInterface(typeof(IDictionary)))
            {
                result = new IDictionaryImplementingFactory<TKey, TValue>(dictionaryType);
                return true;
            }

            // Check if the type has Add(TKey, TValue) such as StringDictionary.
            var addMethod = dictionaryType.GetMethod("Add", [typeof(TKey), typeof(TValue)]);
            if (addMethod != null)
            {
                result = new AddDictionaryFactory<TKey, TValue>(dictionaryType);
                return true;
            }
        }
        else
        {
            // Check if the type has .ctor(IDictionary<TKey, TValue>) such as ReadOnlyDictionary<TKey, TValue>.
            var ctor = dictionaryType.GetConstructor([typeof(IDictionary<TKey, TValue>)]);
            if (ctor != null)
            {
                result = new ConstructorDictionaryFactory<TKey, TValue>(dictionaryType);
                return true;
            }
        }

        result = default;
        return false;
    }

    internal static bool TryCreateObjectMap<T>(FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ExcelClassMap<T>? classMap)
    {
        Type type = typeof(T);
        if (type.GetTypeInfo().IsInterface)
        {
            classMap = null;
            return false;
        }
        if (type.IsAbstract)
        {
            classMap = null;
            return false;
        }
        if (type.GetConstructor([]) is null)
        {
            classMap = null;
            return false;
        }

        var map = new ExcelClassMap<T>(emptyValueStrategy);
        IEnumerable<MemberInfo> properties = type.GetRuntimeProperties().Where(ShouldAutoMapProperty);
        IEnumerable<MemberInfo> fields = type.GetRuntimeFields().Where(ShouldAutoMapField);

        foreach (MemberInfo member in properties.Concat(fields))
        {
            // Ignore this property/field.
            if (Attribute.IsDefined(member, typeof(ExcelIgnoreAttribute)))
            {
                continue;
            }

            // Infer the mapping for each member (property/field) belonging to the type.
            Type memberType = member.MemberType();
            if (memberType == type)
            {
                throw new ExcelMappingException($"Cannot map recursive property \"{member.Name}\" of type {memberType}. Consider applying the ExcelIgnore attribute.");
            }

            var method = TryAutoMapMemberMethod.MakeGenericMethod(memberType);
            var parameters = new object?[] { member, emptyValueStrategy, null };
            var result = (bool)method.InvokeUnwrapped(null, parameters)!;
            if (!result)
            {
                classMap = null;
                return false;
            }

            // Get the out parameter representing the map for the member.
            map.Properties.Add(new ExcelPropertyMap(member, (IMap)parameters[2]!));
        }

        classMap = map;
        return true;
    }

    private static bool ShouldAutoMapProperty(PropertyInfo property)
    {
        // Property must have a setter.
        if (!property.CanWrite)
        {
            return false;
        }
        // Property must be a public instance property.
        if (!property.SetMethod!.IsPublic || property.SetMethod.IsStatic)
        {
            return false;
        }

        // Property must not be an indexer.
        if (property.GetIndexParameters().Length > 0)
        {
            return false;
        }

        // Otherwise, this property can be mapped.
        return true;
    }

    private static bool ShouldAutoMapField(FieldInfo field)
    {
        // Property must be a public instance property.
        if (!field.IsPublic || field.IsStatic)
        {
            return false;
        }

        return true;
    }

    private static FallbackStrategy GetEmptyValueStrategy(IMap map)
    {
        if (map is ExcelClassMap classMap)
        {
            return classMap.EmptyValueStrategy;
        }

        return FallbackStrategy.ThrowIfPrimitive;
    }

    private static IMap? GetExistingMap(IMap parentMap, MemberInfo? member, object? index)
    {
        if (parentMap is ExcelClassMap parentClassMap)
        {
            return parentClassMap.Properties.FirstOrDefault(m => m.Member.Equals(member))?.Map;
        }
        else if (parentMap is IEnumerableIndexerMap enumerableIndexerMap)
        {
            return enumerableIndexerMap.Values.TryGetValue((int)index!, out var map) ? map : null;
        }
        else if (parentMap is IMultidimensionalIndexerMap multidimensionalIndexerMap)
        {
            // Need to find the key that matches the indices.
            // Cannot use TryGetValue as arrays do not implement equality.
            var indices = (int[])index!;
            foreach (var kvp in multidimensionalIndexerMap.Values)
            {
                if (kvp.Key.SequenceEqual(indices))
                {
                    return kvp.Value;
                }
            }

            return null;
        }
        else
        {
            var dictionaryIndexerMap = (IDictionaryIndexerMap)parentMap;
            return dictionaryIndexerMap.Values.TryGetValue((string)index!, out var map) ? map : null;
        }
    }

    private static IMap AddExistingMap(IMap parentMap, MemberInfo? member, object? index, IMap map)
    {
        if (parentMap is ExcelClassMap parentClassMap)
        {
            parentClassMap.Properties.Add(new ExcelPropertyMap(member!, map));
        }
        else if (parentMap is IEnumerableIndexerMap enumerableIndexerMap)
        {
            enumerableIndexerMap.Values[(int)index!] = map;
        }
        else if (parentMap is IMultidimensionalIndexerMap multidimensionalIndexerMap)
        {
            multidimensionalIndexerMap.Values[(int[])index!] = map;
        }
        else
        {
            var dictionaryIndexerMap = (IDictionaryIndexerMap)parentMap;
            dictionaryIndexerMap.Values[(string)index!] = map;
        }

        return map;
    }

    private static MethodInfo? s_getOrCreateNestedMapGenericMethod;
    private static MethodInfo GetOrCreateNestedMapGenericMethod => s_getOrCreateNestedMapGenericMethod ??= typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(GetOrCreateNestedMapGeneric))!;

    public static ExcelClassMap GetOrCreateNestedMap(IMap parentMap, MemberInfo? member, Type memberType, object? index)
    {
        var method = GetOrCreateNestedMapGenericMethod.MakeGenericMethod(memberType);
        var parameters = new object?[] { parentMap, member, index };
        return (ExcelClassMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static IMap GetOrCreateNestedMapGeneric<TProperty>(IMap parentMap, MemberInfo member, object? index)
    {
        if (GetExistingMap(parentMap, member, index) is { } existingMap)
        {
            if (existingMap is not ExcelClassMap<TProperty> existingTypedMap)
            {
                throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
            }

            return existingMap;
        }

        // By default, do not auto map nested fields.
        var classMap = new ExcelClassMap<TProperty>();
        AddExistingMap(parentMap, member, index, classMap);
        return classMap;
    }

    /// <summary>
    /// Creates a class map for the given type using the given strategy.
    /// </summary>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    /// <param name="classMap">The class map for the given type.</param>
    /// <returns>True if the class map could be created, else false.</returns>
    public static bool TryCreateClassMap<T>(FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ExcelClassMap<T>? result)
    {
        if (!Enum.IsDefined(typeof(FallbackStrategy), emptyValueStrategy))
        {
            throw new ArgumentException($"Invalid value \"{emptyValueStrategy}\".", nameof(emptyValueStrategy));
        }

        // Mapping a known type (e.g., string, object, etc.) is supported and simply
        // reads the first column into that value.
        if (TryGetWellKnownMap<T>(emptyValueStrategy, out var mapper, out var emptyFallback, out var invalidFallback))
        {
            var builtinMap = new OneToOneMap<T>(new ColumnIndexReaderFactory(0))
                .WithCellValueMappers(mapper!)
                .WithEmptyFallbackItem(emptyFallback!)
                .WithInvalidFallbackItem(invalidFallback!);
            result = new BuiltinClassMap<T>(builtinMap);
            return true;
        }
        // User may ask to map the row to a dictionary.
        else if (TryCreateDictionaryMap<T>(null, emptyValueStrategy, true, out var dictionaryMap))
        {
            result = new BuiltinClassMap<T>(dictionaryMap);
            return true;
        }
        // User may ask to map the row to a list.
        else if (TryCreateEnumerableMap<T>(emptyValueStrategy, out var enumerableMap))
        {
            result = new BuiltinClassMap<T>(enumerableMap);
            return true;
        }
        // Otherwise, create the default class map for this type.
        else if (TryCreateObjectMap<T>(emptyValueStrategy, out var classMap))
        {
            result = classMap;
            return true;
        }

        result = null;
        return false;
    }

    private static void ApplyMemberAttributesToMap(MemberInfo? member, IToOneMap map)
    {
        // If no member is specified, there is nothing to apply.
        if (member == null)
        {
            return;
        }

        if (Attribute.IsDefined(member, typeof(ExcelOptionalAttribute)))
        {
            map.Optional = true;
        }
        if (Attribute.IsDefined(member, typeof(ExcelPreserveFormattingAttribute)))
        {
            map.PreserveFormatting = true;
        }
    }

    private class BuiltinClassMap<T> : ExcelClassMap<T>
    {
        private IMap BuiltinMap { get; }

        public BuiltinClassMap(IMap primitiveMap)
        {
            BuiltinMap = primitiveMap;
        }

        public override bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
            => BuiltinMap.TryGetValue(sheet, rowIndex, reader, null, out result);
    }
}
