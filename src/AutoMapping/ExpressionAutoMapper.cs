using System.Diagnostics.CodeAnalysis;
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
    private static IMap GetOrCreateMap<T>(ExcelClassMap<T> classMap, Expression expression, Func<MappedExpression, EmptyValueStrategy, IMap> createMap)
    {
        var emptyValueStrategy = classMap.EmptyValueStrategy;
        if (SkipConvert(expression, out var convertedType) is ParameterExpression parameter)
        {
            // The row itself, such as Map(p => p).
            var valueMap = createMap(new MappedParameterExpression(convertedType ?? parameter.Type), emptyValueStrategy);
            classMap.SingleValueMap = valueMap;
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
        var map = createMap(current, emptyValueStrategy);
        ((IParentMap)parentMap).SetChild(current.Key, map);
        return map;
    }

    // The type a collection, dictionary or row number is created as: a member's declared type, so
    // that it can be assigned to the member.
    private static Type GetCreatedType(MappedExpression expression)
        => expression is MappedMemberExpression memberExpression ? memberExpression.Type : expression.MappedValueType;

    public static ValueMap<TProperty> GetOrCreateValueMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap createMap(MappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
            => AutoMapper.CreateValueMap<TProperty>(finalExpression.Member, finalExpression.GetDefaultColumns(), emptyValueStrategy, isAutoMapping: false)!;

        return (ValueMap<TProperty>)GetOrCreateMap(classMap, expression, createMap);
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

        return (CollectionMap<TElement>)GetOrCreateMap(classMap, expression, createMap);
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

        var map = (DictionaryMap<TKey, TValue>)GetOrCreateMap(classMap, expression, createMap);

        // Honour [ExcelExtraColumns] on a member mapped with Map(...), as auto-mapping does.
        if (MapTree.FindOwner(classMap, map, out var member) is { } owner && member is not null && Attribute.IsDefined(member, typeof(ExcelExtraColumnsAttribute)))
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

        return (RowNumberMap)GetOrCreateMap(classMap, expression, createMap);
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

        return (ExcelClassMap<TElement>)GetOrCreateMap(classMap, expression, createMap);
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

        var indexExpression = ExpressionAutoMapper.SkipConvert(expression.Right, out var _);
        if (indexExpression is not ConstantExpression constantExpression)
        {
            throw new ArgumentException($"The indexer must be a constant expression. Received {expression}.", nameof(expression));
        }
        if ((int)constantExpression.Value! < 0)
        {
            throw new ArgumentException($"Array and list indexers must be non-negative. Received: {constantExpression.Value!}", nameof(expression));
        }

        Index = (int)constantExpression.Value!;
    }

    public MappedEnumerableIndexerExpression(MethodCallExpression expression, Type? mappedElementType)
    {
        var indexExpression = ExpressionAutoMapper.SkipConvert(expression.Arguments[0], out var _);

        Type = expression.Object!.Type;
        MappedValueType = mappedElementType ?? Type.GetElementTypeOrEnumerableType()!;

        if (indexExpression is not ConstantExpression constantExpression)
        {
            throw new ArgumentException($"The indexer must be a constant expression. Received {expression}.", nameof(expression));
        }
        if ((int)constantExpression.Value! < 0)
        {
            throw new ArgumentException($"Array and list indexers must be non-negative. Received: {constantExpression.Value!}", nameof(expression));
        }

        Index = (int)constantExpression.Value!;
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
            var indexExpression = ExpressionAutoMapper.SkipConvert(arguments[i], out var _);
            if (indexExpression is not ConstantExpression constantExpression)
            {
                throw new ArgumentException($"Array indices must be constant expressions. Received: {indexExpression}.", nameof(expression));
            }
            if ((int)constantExpression.Value! < 0)
            {
                throw new ArgumentException($"Array indices must be non-negative. Received: {constantExpression.Value!}", nameof(expression));
            }

            Indices[i] = (int)constantExpression.Value!;
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
        var keyExpression = ExpressionAutoMapper.SkipConvert(expression.Arguments[0], out var _);
        if (keyExpression is not ConstantExpression constantExpression)
        {
            throw new ArgumentException($"Dictionary indexer key must be a constant expression. Received: {expression}.", nameof(expression));
        }
        if (constantExpression.Value == null)
        {
            throw new ArgumentException($"Dictionary indexer key cannot be null. Received: {constantExpression.Value!}", nameof(expression));
        }

        Type = expression.Object!.Type;
        DictionaryRule.TryGetKeyValueTypes(Type, out _, out var actualValueType);

        MappedValueType = valueType ?? actualValueType!;

        Key = constantExpression.Value!;
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
