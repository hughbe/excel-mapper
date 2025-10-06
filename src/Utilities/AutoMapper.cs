using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

public static class AutoMapper
{
    private static MethodInfo? s_tryAutoMapMemberMethod;
    private static MethodInfo TryAutoMapMemberMethod => s_tryAutoMapMemberMethod ?? (s_tryAutoMapMemberMethod = typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryAutoMapMember)));

    private static MethodInfo? s_tryCreateSplitGenericMapMethod;
    private static MethodInfo TryCreateSplitGenericMapMethod => s_tryCreateSplitGenericMapMethod ?? (s_tryCreateSplitGenericMapMethod = typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryCreateSplitMapGeneric)));

    private static MethodInfo? s_tryCreateGenericDictionaryMapMethod;
    private static MethodInfo TryCreateGenericDictionaryMapMethod => s_tryCreateGenericDictionaryMapMethod ?? (s_tryCreateGenericDictionaryMapMethod = typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryCreateGenericDictionaryMap)));

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
        if (TryCreateDictionaryMap<TMember>(member, emptyValueStrategy, out var dictionaryMap))
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

    internal static bool TryCreatePrimitivePipeline<T>(FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ValuePipeline<T>? pipeline)
    {
        if (!TryGetWellKnownMap<T>(emptyValueStrategy, out ICellMapper? mapper, out IFallbackItem? emptyFallback, out IFallbackItem? invalidFallback))
        {
            pipeline = null;
            return false;
        }

        pipeline = new ValuePipeline<T>();
        pipeline.AddCellValueMapper(mapper);
        pipeline.EmptyFallback = emptyFallback;
        pipeline.InvalidFallback = invalidFallback;
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
            var matcher = (IExcelColumnMatcher)Activator.CreateInstance(columnNameMatchingAttribute.Type, columnNameMatchingAttribute.ConstructorArguments);
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
            return new ColumnIndicesReaderFactory(colummnIndexAttributes.Select(c => c.Index).ToArray());
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

    internal static bool TryCreateSplitMap(MemberInfo member, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        Type listType = member.MemberType();
        if (!listType.GetElementTypeOrEnumerableType(out Type? elementType))
        {
            map = null;
            return false;
        }

        MethodInfo method = TryCreateSplitGenericMapMethod.MakeGenericMethod([listType, elementType]);

        var parameters = new object?[] { member, emptyValueStrategy, null };
        bool result = (bool)method.Invoke(null, parameters);
        if (result)
        {
            map = (IMap)parameters[2]!;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryCreateSplitMap<TElement>(MemberInfo member, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        MethodInfo method = TryCreateSplitGenericMapMethod.MakeGenericMethod([member.MemberType(), typeof(TElement)]);

        var parameters = new object?[] { member, emptyValueStrategy, null };
        bool result = (bool)method.Invoke(null, parameters);
        if (result)
        {
            map = (IMap)parameters[2]!;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryCreateEnumerableMap<T>(FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        Type listType = typeof(T);
        if (!listType.GetElementTypeOrEnumerableType(out Type? elementType))
        {
            map = null;
            return false;
        }

        MethodInfo method = TryCreateSplitGenericMapMethod.MakeGenericMethod([listType, elementType]);

        var parameters = new object?[] { null, emptyValueStrategy, null };
        bool result = (bool)method.Invoke(null, parameters);
        if (result)
        {
            map = (IMap)parameters[parameters.Length - 1]!;
            return true;
        }

        map = null;
        return false;
    }

    private static bool TryCreateSplitMapGeneric<TList, TElement>(MemberInfo? member, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ManyToOneEnumerableMap<TElement>? map)
    {
        // First, get the pipeline for the element. This is used to convert individual values
        // to be added to/included in the collection.
        if (!TryCreatePrimitivePipeline<TElement>(emptyValueStrategy, out var elementMapping))
        {
            map = null;
            return false;
        }

        // Secondly, find the right way of adding the converted value to the collection.
        if (!TryGetCreateElementsFactory<TList, TElement>(out var factory))
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
            var matcher = (IExcelColumnMatcher)Activator.CreateInstance(columnNameMatchingAttribute.Type, columnNameMatchingAttribute.ConstructorArguments);
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

    private static bool TryGetCreateElementsFactory<TList, TElement>([NotNullWhen(true)] out CreateElementsFactory<TElement>? result)
    {
        Type listType = typeof(TList);
        if (listType.IsArray)
        {
            result = elements => elements.ToArray();
            return true;
        }
        else if (listType.IsImmutableEnumerableType())
        {
            MethodInfo createRangeMethod = listType.GetImmutableEnumerableCreateRangeMethod(typeof(TElement));
            result = elements =>
            {
                return (IEnumerable<TElement>)createRangeMethod.Invoke(null, [elements]);
            };
            return true;
        }
        else if (listType.IsInterface)
        {
            // Add values by creating a list and assigning to the property.
            if (listType.IsAssignableFrom(typeof(List<TElement>).GetTypeInfo()))
            {
                result = elements => elements;
                return true;
            }
        }
        else if (listType.ImplementsInterface(typeof(ICollection<TElement>)))
        {
            result = elements =>
            {
                var value = (ICollection<TElement?>)Activator.CreateInstance(listType);
                foreach (TElement? element in elements)
                {
                    value.Add(element);
                }

                return value;
            };
            return true;
        }

        // Check if the type has .ctor(IEnumerable<T>) such as Queue or Stack.
        ConstructorInfo? ctor = listType.GetConstructor([typeof(IEnumerable<TElement>)]);
        if (ctor != null)
        {
            result = element =>
            {
                return (IEnumerable<TElement?>)Activator.CreateInstance(listType, [element]);
            };
            return true;
        }

        // Check if the type has Add(T) such as BlockingCollection.
        MethodInfo? addMethod = listType.GetMethod("Add", [typeof(TElement)]);
        if (addMethod != null)
        {
            result = elements =>
            {
                var value = Activator.CreateInstance(listType);
                foreach (TElement? element in elements)
                {
                    addMethod.Invoke(value, [element]);
                }

                return value;
            };
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetDictionaryKeyValueType<T>([NotNullWhen(true)] out Type? keyType, [NotNullWhen(true)] out Type? valueType)
    {
        // We should be able to parse anything that implements IEnumerable<KeyValuePair<TKey, TValue>>
        if (typeof(T).ImplementsGenericInterface(typeof(IEnumerable<>), out Type? keyValuePairType))
        {
            if (keyValuePairType.IsGenericType && keyValuePairType.GetGenericTypeDefinition() == typeof(KeyValuePair<,>))
            {
                Type[] arguments = keyValuePairType.GenericTypeArguments;
                keyType = arguments[0];
                valueType = arguments[1];
                return true;
            }
        }

        // Otherwise we can parse regular IDictionary.
        if (typeof(T) == typeof(IDictionary) || typeof(T).ImplementsInterface(typeof(IDictionary)))
        {
            keyType = typeof(string);
            valueType = typeof(object);
            return true;
        }
        
        keyType = null;
        valueType = null;
        return false;
    }

    private static bool TryCreateDictionaryMap<TMember>(MemberInfo? member, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        if (!TryGetDictionaryKeyValueType<TMember>(out var keyType, out var valueType))
        {
            map = null;
            return false;
        }

        MethodInfo method = TryCreateGenericDictionaryMapMethod.MakeGenericMethod(keyType, valueType);
        var parameters = new object?[] { member, typeof(TMember), emptyValueStrategy, null };
        bool result = (bool)method.Invoke(null, parameters);
        if (result)
        {
            map = (IMap)parameters[parameters.Length - 1]!;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryCreateGenericDictionaryMap<TKey, TValue>(MemberInfo? member, Type memberType, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ManyToOneDictionaryMap<TValue>? map)
    {
        if (!TryCreatePrimitivePipeline(emptyValueStrategy, out ValuePipeline<TValue>? valuePipeline))
        {
            map = null;
            return false;
        }

        if (!TryGetCreateDictionaryFactory<TKey, TValue>(memberType, out CreateDictionaryFactory<TValue>? factory))
        {
            map = null;
            return false;
        }

        // Default to all columns.
        var defaultReaderFactory = GetDefaultCellsReaderFactory(member) ?? new AllColumnNamesReaderFactory();
        map = new ManyToOneDictionaryMap<TValue>(defaultReaderFactory, valuePipeline, factory);
        ApplyMemberAttributesToMap(member, map);
        return true;
    }

    private static bool TryGetCreateDictionaryFactory<TKey, TValue>(Type memberType, [NotNullWhen(true)] out CreateDictionaryFactory<TValue>? result)
    {
        if (memberType.IsImmutableDictionaryType())
        {
            MethodInfo createRangeMethod = memberType.GetImmutableDictionaryCreateRangeMethod(typeof(TValue));
            result = elements =>
            {
                return (IDictionary<string, TValue>)createRangeMethod.Invoke(null, [elements]);
            };
            return true;
        }
        if (memberType.GetTypeInfo().IsInterface)
        {
            if (memberType.GetTypeInfo().IsAssignableFrom(typeof(Dictionary<TKey, TValue>).GetTypeInfo()))
            {
                result = elements =>
                {
                    var dictionary = new Dictionary<string, TValue>();
                    foreach (KeyValuePair<string, TValue> keyValuePair in elements)
                    {
                        dictionary.Add(keyValuePair.Key, keyValuePair.Value);
                    }

                    return dictionary;
                };
                return true;
            }
        }
        else if (memberType.ImplementsInterface(typeof(IDictionary<TKey, TValue>)))
        {
            result = elements =>
            {
                IDictionary<string, TValue> dictionary = (IDictionary<string, TValue>)Activator.CreateInstance(memberType);
                foreach (KeyValuePair<string, TValue> keyValuePair in elements)
                {
                    dictionary.Add(keyValuePair);
                }

                return dictionary;
            };
            return true;
        }
        else if (memberType.ImplementsInterface(typeof(IDictionary)))
        {
            result = elements =>
            {
                IDictionary dictionary = (IDictionary)Activator.CreateInstance(memberType);
                foreach (KeyValuePair<string, TValue> keyValuePair in elements)
                {
                    dictionary.Add(keyValuePair.Key, keyValuePair.Value);
                }

                return dictionary;
            };
            return true;
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

        var map = new ExcelClassMap<T>(emptyValueStrategy);
        IEnumerable<MemberInfo> properties = type.GetRuntimeProperties().Where(p => p.CanWrite && p.SetMethod.IsPublic && !p.SetMethod.IsStatic);
        IEnumerable<MemberInfo> fields = type.GetRuntimeFields().Where(f => f.IsPublic && !f.IsStatic);

        foreach (MemberInfo member in properties.Concat(fields))
        {
            // Ignore this property/field.
            if (Attribute.IsDefined(member, typeof(ExcelIgnoreAttribute)))
            {
                continue;
            }

            // Infer the mapping for each member (property/field) belonging to the type.
            Type memberType = member.MemberType();
            MethodInfo method = TryAutoMapMemberMethod.MakeGenericMethod(memberType);
            if (memberType == type)
            {
                throw new ExcelMappingException($"Cannot map recursive property \"{member.Name}\" of type {memberType}. Consider applying the ExcelIgnore attribute.");
            }

            var parameters = new object?[] { member, emptyValueStrategy, null };
            bool result = (bool)method.Invoke(null, parameters);
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
        else if (TryCreateDictionaryMap<T>(null, emptyValueStrategy, out var dictionaryMap))
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
