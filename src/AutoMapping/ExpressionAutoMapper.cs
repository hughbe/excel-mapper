using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Linq;
using System.Linq.Expressions;
using ExcelMapper.Readers;

namespace ExcelMapper.AutoMapping;

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal static class ExpressionAutoMapper
{
    // The member accesses and indexers of an expression such as p => p.Address.Street or
    // p => p.Values[0], with the outermost (the one on the row) on top.
    private static Stack<MappedAccessExpression> BuildAccessStack(Expression expression)
    {
        var expressions = new Stack<MappedAccessExpression>();
        var currentExpression = expression;
        while (true)
        {
            currentExpression = SkipConvert(currentExpression, out var currentTargetType);
            if (currentExpression is ParameterExpression)
            {
                return expressions;
            }
            else if (currentExpression is MemberExpression memberExpression)
            {
                currentExpression = memberExpression.Expression!;
                expressions.Push(new MappedMemberExpression(memberExpression, currentTargetType));
            }
            else if (IsDictionaryIndexerExpression(currentExpression))
            {
                var methodCallExpression = (MethodCallExpression)currentExpression;
                currentExpression = SkipConvert(methodCallExpression.Object!, out var _);
                expressions.Push(new MappedDictionaryIndexerExpression(methodCallExpression, currentTargetType));
            }
            else if (IsArrayIndexerExpression(currentExpression))
            {
                var binaryExpression = (BinaryExpression)currentExpression;
                currentExpression = SkipConvert(binaryExpression.Left, out var _);
                expressions.Push(new MappedEnumerableIndexerExpression(binaryExpression, currentTargetType));
            }
            else if (IsMultidimensionalArrayIndexerExpression(currentExpression))
            {
                var methodCallExpression = (MethodCallExpression)currentExpression;
                currentExpression = SkipConvert(methodCallExpression.Object!, out var _);
                expressions.Push(new MappedMultidimensionalArrayIndexerExpression(methodCallExpression, currentTargetType));
            }
            else if (IsListIndexerExpression(currentExpression))
            {
                var methodCallExpression = (MethodCallExpression)currentExpression;
                currentExpression = SkipConvert(methodCallExpression.Object!, out var _);
                expressions.Push(new MappedEnumerableIndexerExpression(methodCallExpression, currentTargetType));
            }
            else
            {
                throw new ArgumentException($"Expression can only contain member accesses, but found {currentExpression.NodeType} ({currentExpression}) of type {currentExpression.GetType()}", nameof(expression));
            }
        }
    }

    internal static Expression SkipConvert(Expression expression, out Type convertedType)
    {
        convertedType = null!;
        while (expression is UnaryExpression unaryExpression && 
               (unaryExpression.NodeType == ExpressionType.Convert || unaryExpression.NodeType == ExpressionType.ConvertChecked))
        {
            expression = unaryExpression.Operand;
            // If we have multiple converts, we want the outermost type
            // e.g. ((A)(B)x) should give us A, not B
            convertedType ??= unaryExpression.Type;
        }

        return expression;
    }
    
    /// <summary>
    /// Gets the value of an index or dictionary key in a mapping expression, as
    /// <paramref name="type"/>: a constant, such as <c>p => p.Values[0]</c>, or a variable or
    /// field the expression captures, such as <c>i</c> in <c>for (var i = 0; ...) Map(p => p.Values[i])</c>.
    /// Anything else, such as a method call or a member of the row, isn't allowed: the value must be
    /// known when the class map is created.
    /// </summary>
    internal static bool TryGetIndexValue(Expression expression, Type type, out object? value)
    {
        if (!TryEvaluate(expression, out value))
        {
            return false;
        }

        // The value of a variable converted in the expression, such as a short index.
        if (value is not null && value.GetType() != type && value is IConvertible && (Nullable.GetUnderlyingType(type) ?? type) is { IsPrimitive: true } primitiveType)
        {
            value = Convert.ChangeType(value, primitiveType, CultureInfo.InvariantCulture);
        }

        return true;
    }

    private static bool TryEvaluate(Expression expression, out object? value)
    {
        switch (SkipConvert(expression, out _))
        {
            case ConstantExpression constant:
                value = constant.Value;
                return true;

            // A captured variable is a field of the closure: a constant. A static field, or a field
            // or property of a captured object, is read now.
            case MemberExpression member:
                object? target = null;
                if (member.Expression is not null && !TryEvaluate(member.Expression, out target))
                {
                    break;
                }

                if (member.Member is FieldInfo field)
                {
                    value = field.GetValue(target);
                    return true;
                }

                if (member.Member is PropertyInfo { GetMethod: not null } property && property.GetIndexParameters().Length == 0)
                {
                    value = property.GetValue(target);
                    return true;
                }

                break;
        }

        value = null;
        return false;
    }

    private static bool IsArrayIndexerExpression(Expression expression)
        => expression is BinaryExpression binaryExpression && binaryExpression.NodeType == ExpressionType.ArrayIndex;

    private static bool IsMultidimensionalArrayIndexerExpression(Expression expression)
        => expression is MethodCallExpression methodCallExpression &&
            methodCallExpression.Method.Name == "Get" &&
            methodCallExpression.Object is not null &&
            methodCallExpression.Object!.Type!.IsArray &&
            methodCallExpression.Arguments.Count > 1 &&
            methodCallExpression.Arguments.All(arg => arg.Type == typeof(int));

    private static bool IsListIndexerExpression(Expression expression)
        => expression is MethodCallExpression methodCallExpression &&
            methodCallExpression.Method.Name == "get_Item" &&
            methodCallExpression.Object is not null &&
            methodCallExpression.Object!.Type!.GetElementTypeOrEnumerableType() != null &&
            methodCallExpression.Arguments.Count == 1 &&
            methodCallExpression.Arguments[0].Type == typeof(int);

    private static bool IsDictionaryIndexerExpression(Expression expression)
        => expression is MethodCallExpression methodCallExpression &&
            methodCallExpression.Object is not null &&
            DictionaryRule.TryGetKeyValueTypes(methodCallExpression.Object!.Type!, out var _, out var _) &&
            methodCallExpression.Method.Name == "get_Item" &&
            methodCallExpression.Arguments.Count == 1;

    /// <summary>
    /// Gets the map of what an expression reads, creating it with <paramref name="createMap"/> if
    /// it's not mapped yet, along with the maps of the objects, collections and dictionaries it's
    /// read through: for <c>p => p.Address.Street</c>, the class map of <c>Address</c>.
    /// </summary>
    /// <remarks>
    /// A member already mapped with <c>Map(...)</c> to the same kind of map keeps its map, so that
    /// <c>Map(p => p.Name).WithColumnName("Full Name")</c> followed by <c>Map(p => p.Name).MakeOptional()</c>
    /// configures one map. A member that auto-mapping mapped (with <c>AutoMap()</c>), or that was
    /// mapped to another kind of map, gets a new one.
    /// </remarks>
    private static TMap GetOrCreateMap<T, TMap>(ExcelClassMap<T> classMap, Expression expression, Func<MappedExpression, EmptyValueStrategy, IMap> createMap, out bool created) where TMap : class, IMap
    {
        // As when a member's map is replaced, a class map that has read rows can't change.
        ThrowHelpers.ThrowIfFrozen(classMap.IsFrozen);

        var emptyValueStrategy = classMap.EmptyValueStrategy;
        if (SkipConvert(expression, out var convertedType) is ParameterExpression parameter)
        {
            // The row itself, such as Map(p => p).
            var valueMap = (TMap)createMap(new MappedParameterExpression(convertedType ?? parameter.Type), emptyValueStrategy);
            classMap.SingleValueMap = valueMap;
            created = true;
            return valueMap;
        }

        var stack = BuildAccessStack(expression);
        IMap parentMap = classMap;
        var current = stack.Pop();
        while (stack.Count > 0)
        {
            var next = stack.Pop();
            parentMap = next.GetOrCreateSourceMap(parentMap, current, emptyValueStrategy);
            current = next;
        }

        if (parentMap is ExcelClassMap && current is not MappedMemberExpression)
        {
            throw new ArgumentException($"Expected a member expression when adding to an object map. Received: {expression}", nameof(expression));
        }

        // A member of a class map, or an element mapped by its index or key.
        var parent = (IParentMap)parentMap;
        if (parent.GetMappedChild(current.Key) is TMap existing)
        {
            created = false;
            return existing;
        }

        var map = (TMap)createMap(current, emptyValueStrategy);
        parent.SetChild(current.Key, map);
        created = true;
        return map;
    }

    // The type a collection, dictionary or row number is created as: a member's declared type, so
    // that it can be assigned to the member.
    private static Type GetCreatedType(MappedExpression expression)
        => expression is MappedMemberExpression memberExpression ? memberExpression.Type : expression.MappedValueType;

    public static ValueMap<TProperty> GetOrCreateValueMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap createMap(MappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            var map = AutoMapper.CreateValueMap<TProperty>(finalExpression.Member, finalExpression.GetDefaultColumns(), emptyValueStrategy, isAutoMapping: false)!;
            if (!map.Pipeline.HasMappers)
            {
                map.Pipeline.NoConversion = DescribeNoConversion(typeof(TProperty), TypeRule<TProperty>.Instance);
            }

            return map;
        }

        return GetOrCreateMap<T, ValueMap<TProperty>>(classMap, expression, createMap, out _);
    }

    // Why a member mapped with Map(...) can't be read from a cell, and what to do instead: the
    // collection and dictionary overloads of Map only cover common types.
    private static string DescribeNoConversion(Type type, TypeRule rule)
    {
        var typeName = TypeNames.Get(type);
        if (rule.Dictionary is { } dictionary)
        {
            return $"of type {typeName} is a dictionary that Map can't read from a cell: map it with MapDictionary<{TypeNames.Get(dictionary.ValueType)}>(...).";
        }

        if (rule.Collection is { } collection)
        {
            return $"of type {typeName} is a collection that Map can't read from a cell: map it with MapCollection<{TypeNames.Get(collection.ElementType)}>(...).";
        }

        return $"of type {typeName} has no conversion from a cell. Give it one with WithConverter or [ExcelConverter], or map its own members with MapObject(...).";
    }

    public static CollectionMap<TElement> GetOrCreateCollectionMap<T, TElement>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap createMap(MappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            var member = finalExpression.Member;
            var type = GetCreatedType(finalExpression);
            var columns = finalExpression.GetDefaultCellsColumns() ?? MemberMapper.GetDefaultSplitColumns(member, finalExpression.GetDefaultColumns());
            if (!CollectionRule<TElement>.TryCreateCollectionMap(member, type, columns, emptyValueStrategy, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type {TypeNames.Get(type)}. It must be a single dimensional array, be assignable from List<T> or implement ICollection<T>.");
            }

            return map;
        }

        return GetOrCreateMap<T, CollectionMap<TElement>>(classMap, expression, createMap, out _);
    }

    public static DictionaryMap<TKey, TValue> GetOrCreateDictionaryMap<T, TKey, TValue>(ExcelClassMap<T> classMap, Expression expression) where TKey : notnull
    {
        static IMap createMap(MappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            var type = GetCreatedType(finalExpression);
            if (!DictionaryRule<TKey, TValue>.TryCreateDictionaryMap(finalExpression.Member, type, emptyValueStrategy, isAutoMapping: false, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type {TypeNames.Get(type)}.");
            }

            return map;
        }

        var map = GetOrCreateMap<T, DictionaryMap<TKey, TValue>>(classMap, expression, createMap, out var created);

        // Honour [ExcelExtraColumns] on a member mapped with Map(...), as auto-mapping does: when
        // it's first mapped, so that columns chosen afterwards stay chosen.
        if (created && MapTree.FindOwner(classMap, map, out var member) is { } owner && member is not null && Attribute.IsDefined(member, typeof(ExcelExtraColumnsAttribute)))
        {
            map.Columns = new ExtraColumns(owner);
        }

        return map;
    }

    public static DictionaryMap<string, TValue> GetOrCreateExtraColumnsMap<T, TValue>(ExcelClassMap<T> classMap, Expression expression)
    {
        var map = GetOrCreateDictionaryMap<T, string, TValue>(classMap, expression);
        var owner = MapTree.FindOwner(classMap, map, out _) ?? throw new ArgumentException($"Extra columns can only be mapped to a property or field. Received: {expression}", nameof(expression));
        map.Columns = new ExtraColumns(owner);
        return map;
    }

    public static RowNumberMap GetOrCreateRowNumberMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap createMap(MappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            var type = GetCreatedType(finalExpression);
            if (!RowNumberMap.IsSupportedType(type))
            {
                throw new ArgumentException($"Row numbers can't be mapped to type {TypeNames.Get(type)}. Use a numeric type such as int or long.", nameof(expression));
            }

            return new RowNumberMap(type);
        }

        return GetOrCreateMap<T, RowNumberMap>(classMap, expression, createMap, out _);
    }

    public static ExcelClassMap<TElement> GetOrCreateObjectMap<T, TElement>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap createMap(MappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            if (!AutoMapper.TryCreateObjectMap<TElement>(emptyValueStrategy, out var map))
            {
                throw new ExcelMappingException($"Could not map object of type {TypeNames.Get(typeof(TElement))}.");
            }

            return map;
        }

        return GetOrCreateMap<T, ExcelClassMap<TElement>>(classMap, expression, createMap, out _);
    }
}

/// <summary>
/// What a mapping expression reads: the row itself (<c>p => p</c>), or a member, element or value
/// (see <see cref="MappedAccessExpression"/>).
/// </summary>
[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal abstract class MappedExpression
{
    /// <summary>
    /// The type the value is mapped as: its type, or the type the expression converts it to.
    /// </summary>
    public abstract Type MappedValueType { get; }

    /// <summary>
    /// The property or field read, if any.
    /// </summary>
    public virtual MemberInfo? Member => null;

    /// <summary>
    /// The columns a single value is read from by default.
    /// </summary>
    public abstract Columns GetDefaultColumns();

    /// <summary>
    /// The columns a collection or dictionary is read from by default, or null to read it from
    /// the cell of <see cref="GetDefaultColumns"/>, split.
    /// </summary>
    public virtual Columns? GetDefaultCellsColumns() => null;
}

/// <summary>
/// The row itself, as in <c>Map(p => p)</c>.
/// </summary>
[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal sealed class MappedParameterExpression(Type mappedValueType) : MappedExpression
{
    public override Type MappedValueType { get; } = mappedValueType;

    public override Columns GetDefaultColumns() => AutoMapper.s_firstColumn;

    public override Columns? GetDefaultCellsColumns() => AllColumns.Instance;
}

/// <summary>
/// A member access or indexer in a mapping expression, such as <c>.Address</c> or <c>[0]</c>.
/// </summary>
[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal abstract class MappedAccessExpression : MappedExpression
{
    /// <summary>
    /// The type of the object, collection or dictionary this reads from.
    /// </summary>
    public Type Type { get; private protected set; } = null!;

    /// <summary>
    /// What this reads, as the key of its map in the map of what it reads from
    /// (<see cref="IParentMap.SetChild"/>): the member, index, indices or dictionary key.
    /// </summary>
    public abstract object Key { get; }

    /// <summary>
    /// Gets or creates the map of what this reads from, as the child of
    /// <paramref name="parentMap"/> for <paramref name="source"/>: for <c>.Street</c> in
    /// <c>p.Address.Street</c>, the class map of <c>Address</c>.
    /// </summary>
    public abstract IMap GetOrCreateSourceMap(IMap parentMap, MappedAccessExpression source, EmptyValueStrategy emptyValueStrategy);
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal sealed class MappedMemberExpression : MappedAccessExpression
{
    public override MemberInfo Member { get; }

    public override Type MappedValueType { get; }

    public override object Key => Member;

    public MappedMemberExpression(MemberExpression expression, Type memberType)
    {
        if (expression.Expression is null)
        {
            throw new ArgumentException($"Static members are not supported. Received: {expression}", nameof(expression));
        }

        Member = expression.Member;
        Type = Member.MemberType();
        MappedValueType = memberType ?? expression.Type;
    }

    public override IMap GetOrCreateSourceMap(IMap parentMap, MappedAccessExpression source, EmptyValueStrategy emptyValueStrategy)
        => TypeRule.For(source.MappedValueType).GetOrCreateNestedMap(parentMap, source.Key, emptyValueStrategy);

    public override Columns GetDefaultColumns() => MemberMapper.GetDefaultColumns(Member);

    public override Columns? GetDefaultCellsColumns() => MemberMapper.GetDefaultCellsColumns(Member);
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal sealed class MappedEnumerableIndexerExpression : MappedAccessExpression
{
    public int Index { get; }

    public override Type MappedValueType { get; }

    public override object Key => Index;

    public MappedEnumerableIndexerExpression(BinaryExpression expression, Type? mappedElementType)
    {
        Type = expression.Left.Type!;
        MappedValueType = mappedElementType ?? Type.GetElementType()!;

        Index = GetIndex(expression.Right, expression);
    }

    public MappedEnumerableIndexerExpression(MethodCallExpression expression, Type? mappedElementType)
    {
        Type = expression.Object!.Type;
        MappedValueType = mappedElementType ?? Type.GetElementTypeOrEnumerableType()!;
        Index = GetIndex(expression.Arguments[0], expression);
    }

    // An array or list index: a constant, or a variable the expression captures.
    internal static int GetIndex(Expression indexExpression, Expression expression)
    {
        if (!ExpressionAutoMapper.TryGetIndexValue(indexExpression, typeof(int), out var value) || value is not int index)
        {
            throw new ArgumentException($"An index must be a constant, or a variable or field captured by the expression, rather than computed from the row or a method call. Received: {expression}.", nameof(expression));
        }

        if (index < 0)
        {
            throw new ArgumentException($"Array and list indices must be non-negative. Received: {index} in {expression}.", nameof(expression));
        }

        return index;
    }

    public override IMap GetOrCreateSourceMap(IMap parentMap, MappedAccessExpression source, EmptyValueStrategy emptyValueStrategy)
        => TypeRule.For(Type).Collection!.GetOrCreateIndexerMap(parentMap, Type, source.Key);

    public override Columns GetDefaultColumns() => new ColumnIndexColumns(Index);
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal sealed class MappedMultidimensionalArrayIndexerExpression : MappedAccessExpression
{
    public int[] Indices { get; }

    public override Type MappedValueType { get; }

    public override object Key => Indices;

    public MappedMultidimensionalArrayIndexerExpression(MethodCallExpression expression, Type? mappedElementType)
    {
        Type = expression.Object!.Type;
        MappedValueType = mappedElementType ?? Type.GetElementType()!;

        var arguments = expression.Arguments;
        Indices = new int[arguments.Count];
        for (int i = 0; i < arguments.Count; i++)
        {
            Indices[i] = MappedEnumerableIndexerExpression.GetIndex(arguments[i], expression);
        }
    }

    public override IMap GetOrCreateSourceMap(IMap parentMap, MappedAccessExpression source, EmptyValueStrategy emptyValueStrategy)
        => TypeRule.For(Type).Collection!.GetOrCreateMultidimensionalIndexerMap(parentMap, source.Key);

    public override Columns GetDefaultColumns() => AutoMapper.s_firstColumn;
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal sealed class MappedDictionaryIndexerExpression : MappedAccessExpression
{
    public override object Key { get; }

    public override Type MappedValueType { get; }

    public MappedDictionaryIndexerExpression(MethodCallExpression expression, Type valueType)
    {
        var keyArgument = expression.Arguments[0];
        if (!ExpressionAutoMapper.TryGetIndexValue(keyArgument, keyArgument.Type, out var key))
        {
            throw new ArgumentException($"A dictionary key must be a constant, or a variable or field captured by the expression, rather than computed from the row or a method call. Received: {expression}.", nameof(expression));
        }
        if (key is null)
        {
            throw new ArgumentException($"A dictionary key can't be null. Received: {expression}.", nameof(expression));
        }

        Type = expression.Object!.Type;
        DictionaryRule.TryGetKeyValueTypes(Type, out _, out var actualValueType);

        MappedValueType = valueType ?? actualValueType!;

        Key = key;
    }

    public override IMap GetOrCreateSourceMap(IMap parentMap, MappedAccessExpression source, EmptyValueStrategy emptyValueStrategy)
        => TypeRule.For(Type).Dictionary!.GetOrCreateIndexerMap(parentMap, Type, source.Key);

    public override Columns GetDefaultColumns()
    {
        if (Key is string keyString)
        {
            if (keyString.Length == 0)
            {
                return AutoMapper.s_firstColumn;
            }

            return new ColumnNameColumns(keyString);
        }
        else if (Key is int keyInt)
        {
            return new ColumnIndexColumns(keyInt);
        }

        return new ColumnNameColumns(Key.ToString()!);
    }
}
