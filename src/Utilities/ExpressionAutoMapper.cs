using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Linq;
using System.Linq.Expressions;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal static class ExpressionAutoMapper
{
    private static Stack<IMappedExpression> BuildExpressionStack(Expression expression)
    {
        var expressions = new Stack<IMappedExpression>();
        var currentExpression = expression;

        while (true)
        {
            currentExpression = SkipConvert(currentExpression, out var currentTargetType);
            if (currentExpression is ParameterExpression parameterExpression)
            {
                // If we have reached the parameter expression and there are no other expressions, we
                // need to map the parameter itself.
                if (expressions.Count == 0)
                {
                    expressions.Push(new MappedParameterExpression(parameterExpression, currentTargetType));
                    return expressions;
                }

                break;
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

        return expressions;
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

    private static void AddMap(IMap parentMap, IMappedExpression expression, IMap map)
    {
        if (parentMap is ExcelClassMap classMap && expression is not MappedMemberExpression)
        {
            if (expression is not MappedParameterExpression)
            {
                throw new ArgumentException($"Expected a member expression when adding to an object map. Received: {expression}", nameof(expression));
            }

            // The row itself, such as Map(p => p).
            classMap._valueMap = map;
            return;
        }

        // A member of a class map, or an element mapped by its index or key.
        ((IParentMap)parentMap).SetChild(expression.Context, map);
    }

    [ExcludeFromCodeCoverage]
    private static IMap ProcessExpression<T, TProperty>(Stack<IMappedExpression> stack, IMap currentMap, IMappedExpression currentExpression, Func<IMappedExpression, EmptyValueStrategy, IMap> memberMapCreator, EmptyValueStrategy emptyValueStrategy)
    {
        if (stack.Count == 0)
        {
            var memberMap = memberMapCreator(currentExpression, emptyValueStrategy);
            AddMap(currentMap, currentExpression, memberMap);
            return memberMap;
        }

        var nextExpression = stack.Pop();
        var nextMap = nextExpression.GetNextMap(currentMap, currentExpression, emptyValueStrategy);
        return ProcessExpression<T, TProperty>(stack, nextMap, nextExpression, memberMapCreator, emptyValueStrategy);
    }

    private static IMap GetOrCreateMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression, Func<IMappedExpression, EmptyValueStrategy, IMap> memberMapCreator)
    {
        var stack = BuildExpressionStack(expression);
        return ProcessExpression<T, TProperty>(stack, classMap, stack.Pop(), memberMapCreator, classMap.EmptyValueStrategy);
    }

    public static ValueMap<TProperty> GetOrCreateValueMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(IMappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            var member = finalExpression is MappedMemberExpression memberExpression ? memberExpression.Member : null;
            return AutoMapper.CreateValueMap<TProperty>(member, finalExpression.GetDefaultColumns(), emptyValueStrategy, isAutoMapping: false)!;
        }

        return (ValueMap<TProperty>)GetOrCreateMap<T, TProperty>(classMap, expression, memberMapCreator);
    }

    public static CollectionMap<TElement> GetOrCreateCollectionMap<T, TElement>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(IMappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            MemberInfo? member;
            Type type;
            if (finalExpression is MappedMemberExpression memberExpression)
            {
                member = memberExpression.Member;
                type = memberExpression.Type;
            }
            else
            {
                member = null;
                type = finalExpression.MappedValueType;
            }
            if (!CollectionRule<TElement>.TryCreateCollectionMap(member, type, finalExpression.GetDefaultCellsColumns() ?? MemberMapper.GetDefaultSplitColumns(member, finalExpression.GetDefaultColumns()), emptyValueStrategy, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type \"{type}\". It must be a single dimensional array, be assignable from List<T> or implement ICollection<T>.");
            }

            return map;
        }
        
        return (CollectionMap<TElement>)GetOrCreateMap<T, TElement>(classMap, expression, memberMapCreator);
    }

    public static DictionaryMap<TKey, TValue> GetOrCreateDictionaryMap<T, TKey, TValue>(ExcelClassMap<T> classMap, Expression expression) where TKey : notnull
    {
        static IMap memberMapCreator(IMappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            MemberInfo? member;
            Type type;
            if (finalExpression is MappedMemberExpression memberExpression)
            {
                member = memberExpression.Member;
                type = memberExpression.Type;
            }
            else
            {
                member = null;
                type = finalExpression.MappedValueType;
            }
            if (!DictionaryRule<TKey, TValue>.TryCreateDictionaryMap(member, type, emptyValueStrategy, isAutoMapping: false, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type \"{type}\".");
            }

            return map;
        }

        var map = (DictionaryMap<TKey, TValue>)GetOrCreateMap<T, TValue>(classMap, expression, memberMapCreator);

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
        static IMap memberMapCreator(IMappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            var type = finalExpression is MappedMemberExpression memberExpression ? memberExpression.Type : finalExpression.MappedValueType;
            if (!RowNumberMap.IsSupportedType(type))
            {
                throw new ArgumentException($"Row numbers can't be mapped to type \"{type}\". Use a numeric type such as int or long.", nameof(expression));
            }

            return new RowNumberMap(type);
        }

        return (RowNumberMap)GetOrCreateMap<T, TProperty>(classMap, expression, memberMapCreator);
    }

    public static ExcelClassMap<TElement> GetOrCreateObjectMap<T, TElement>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(IMappedExpression finalExpression, EmptyValueStrategy emptyValueStrategy)
        {
            if (!AutoMapper.TryCreateObjectMap<TElement>(emptyValueStrategy, out var map))
            {
                throw new ExcelMappingException($"Could not map object of type \"{typeof(TElement)}\".");
            }

            return map;
        }

        return (ExcelClassMap<TElement>)GetOrCreateMap<T, TElement>(classMap, expression, memberMapCreator);
    }
}

internal interface IMappedExpression
{
    Type Type { get; }
    Type MappedValueType { get; }
    object Context { get; }
    IMap GetNextMap(IMap currentMap, IMappedExpression currentExpression, EmptyValueStrategy emptyValueStrategy);
    Columns GetDefaultColumns();
    Columns? GetDefaultCellsColumns();
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal class MappedParameterExpression : IMappedExpression
{
    [ExcludeFromCodeCoverage]
    public Type Type { get; }
    public Type MappedValueType { get; }
    [ExcludeFromCodeCoverage]
    public object Context => throw new NotImplementedException();

    public MappedParameterExpression(ParameterExpression expression, Type mappedType)
    {
        Type = expression.Type;
        MappedValueType = mappedType ?? expression.Type;
    }

    [ExcludeFromCodeCoverage]
    public IMap GetNextMap(IMap currentMap, IMappedExpression currentExpression, EmptyValueStrategy emptyValueStrategy)
        => throw new NotImplementedException();

    public Columns GetDefaultColumns() => AutoMapper.s_firstColumn;

    public Columns? GetDefaultCellsColumns() => AllColumns.Instance;
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal class MappedMemberExpression : IMappedExpression
{
    public Type Type { get; }
    public MemberInfo Member { get; }
    public Type MappedValueType { get; }
    public object Context => Member;

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

    public IMap GetNextMap(IMap currentMap, IMappedExpression currentExpression, EmptyValueStrategy emptyValueStrategy)
        => TypeRule.For(currentExpression.MappedValueType).GetOrCreateNestedMap(currentMap, currentExpression.Context, emptyValueStrategy);

    public Columns GetDefaultColumns()
        => MemberMapper.GetDefaultColumns(Member);

    public Columns? GetDefaultCellsColumns() => MemberMapper.GetDefaultCellsColumns(Member);
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal class MappedEnumerableIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type MappedElementType { get; }
    public int Index { get; }
    public Type MappedValueType => MappedElementType;
    public object Context => Index;

    public MappedEnumerableIndexerExpression(BinaryExpression expression, Type? mappedElementType)
    {
        Type = expression.Left.Type!;
        MappedElementType = mappedElementType ?? Type.GetElementType()!;

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
        MappedElementType = mappedElementType ?? Type.GetElementTypeOrEnumerableType()!;

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

    public IMap GetNextMap(IMap currentMap, IMappedExpression currentExpression, EmptyValueStrategy emptyValueStrategy)
        => TypeRule.For(Type).Collection!.GetOrCreateIndexerMap(currentMap, Type, currentExpression.Context);

    public Columns GetDefaultColumns()
        => new ColumnIndexColumns(Index);

    public Columns? GetDefaultCellsColumns() => null;
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal class MappedMultidimensionalArrayIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type MappedElementType { get; }
    public int[] Indices { get; }
    public Type MappedValueType => MappedElementType;
    public object Context => Indices;

    public MappedMultidimensionalArrayIndexerExpression(MethodCallExpression expression, Type? mappedElementType)
    {
        Type = expression.Object!.Type;
        MappedElementType = mappedElementType ?? Type.GetElementType()!;

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

    public IMap GetNextMap(IMap currentMap, IMappedExpression currentExpression, EmptyValueStrategy emptyValueStrategy)
        => TypeRule.For(Type).Collection!.GetOrCreateMultidimensionalIndexerMap(currentMap, currentExpression.Context);

    public Columns GetDefaultColumns()
        => AutoMapper.s_firstColumn;

    public Columns? GetDefaultCellsColumns() => null;
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal class MappedDictionaryIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type ValueType { get; }
    public object Key { get; }
    public Type MappedValueType => ValueType;
    public object Context => Key;

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

        ValueType = valueType ?? actualValueType!;

        Key = constantExpression.Value!;
    }

    public IMap GetNextMap(IMap currentMap, IMappedExpression currentExpression, EmptyValueStrategy emptyValueStrategy)
        => TypeRule.For(Type).Dictionary!.GetOrCreateIndexerMap(currentMap, Type, currentExpression.Context);

    public Columns GetDefaultColumns()
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

    public Columns? GetDefaultCellsColumns() => null;
}
