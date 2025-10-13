using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Linq;
using System.Linq.Expressions;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

internal static class ExpressionAutoMapper
{
    internal static Expression SkipConvert(Expression expression, out Type convertedType)
    {
        convertedType = null!;
        while (expression is UnaryExpression unaryExpression && 
               (unaryExpression.NodeType == ExpressionType.Convert || unaryExpression.NodeType == ExpressionType.ConvertChecked))
        {
            expression = unaryExpression.Operand;
            // If we have multiple converts, we want the innermost type
            // e.g. ((A)(B)x) should give us A, not B
            convertedType ??= unaryExpression.Type;
        }

        return expression;
    }

    private static Stack<IMappedExpression> BuildExpressionStack(Expression expression)
    {
        var expressions = new Stack<IMappedExpression>();
        var currentExpression = expression;
        
        while (true)
        {
            currentExpression = SkipConvert(currentExpression, out var currentTargetType);
            if (currentExpression is MemberExpression memberExpression)
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
            else if (currentExpression is ParameterExpression)
            {
                break;
            }
            else
            {
                throw new ArgumentException($"Expression can only contain member accesses, but found {currentExpression.NodeType} ({currentExpression}) of type {currentExpression.GetType()}", nameof(expression));
            }
        }
        
        return expressions;
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
            AutoMapper.TryGetDictionaryKeyValueType(methodCallExpression.Object!.Type!, out var _, out var _) &&
            methodCallExpression.Method.Name == "get_Item" &&
            methodCallExpression.Arguments.Count == 1;

    private static void AddMap(IMap parentMap, IMappedExpression expression, IMap map)
    {
        if (parentMap is ExcelClassMap classMap)
        {
            if (expression is not MappedMemberExpression memberExpression)
            {
                throw new ArgumentException($"Expected a member expression when adding to an object map. Received: {expression}", nameof(expression));
            }

            var member = memberExpression.Member;
            var existingPropertyMap = classMap.Properties.FirstOrDefault(m => m.Member.Equals(member));
            if (existingPropertyMap is not null)
            {
                classMap.Properties.Remove(existingPropertyMap);
            }

            classMap.Properties.Add(new ExcelPropertyMap(member, map));
        }
        else if (parentMap is IEnumerableIndexerMap arrayIndexerMap)
        {
            var index = ((MappedEnumerableIndexerExpression)expression).Index;
            arrayIndexerMap.Values[index] = map;
        }
        else  if (parentMap is IMultidimensionalIndexerMap multidimensionalArrayIndexerMap)
        {
            var indices = ((MappedMultidimensionalArrayIndexerExpression)expression).Indices;

            // Need to find the key that matches the indices.
            // Cannot use TryGetValue as arrays do not implement equality.
            var key = multidimensionalArrayIndexerMap.Values.Keys.FirstOrDefault(k => k.SequenceEqual(indices)) ?? indices;
            multidimensionalArrayIndexerMap.Values[key] = map;
        }
        else
        {
            var dictionaryIndexerMap = (IDictionaryIndexerMap)parentMap;
            var key = ((MappedDictionaryIndexerExpression)expression).Key;
            dictionaryIndexerMap.Values[key] = map;
        }
    }

    [ExcludeFromCodeCoverage]
    private static IMap ProcessExpression<T, TProperty>(Stack<IMappedExpression> stack, IMap currentMap, IMappedExpression currentExpression, Func<IMappedExpression, FallbackStrategy, IMap> memberMapCreator, FallbackStrategy emptyValueStrategy)
    {
        if (stack.Count == 0)
        {
            var memberMap = memberMapCreator(currentExpression, emptyValueStrategy);
            AddMap(currentMap, currentExpression, memberMap);
            return memberMap;
        }

        var nextExpression = stack.Pop();
        var nextMap = GetNextMap(currentMap, currentExpression, nextExpression, emptyValueStrategy);
        return ProcessExpression<T, TProperty>(stack, nextMap, nextExpression, memberMapCreator, emptyValueStrategy);
    }

    [ExcludeFromCodeCoverage]
    private static IMap GetNextMap(IMap currentMap, IMappedExpression currentExpression, IMappedExpression nextExpression, FallbackStrategy emptyValueStrategy)
    {
        return nextExpression switch
        {
            MappedMemberExpression memberExpression => AutoMapper.GetOrCreateNestedMap(currentMap, currentExpression.MappedValueType, currentExpression.Context, emptyValueStrategy),
            MappedDictionaryIndexerExpression dictionaryIndexerExpression => AutoMapper.GetOrCreateDictionaryIndexerMap(currentMap, dictionaryIndexerExpression.Type, currentExpression.Context, dictionaryIndexerExpression.ActualKeyType, dictionaryIndexerExpression.ActualValueType),
            MappedEnumerableIndexerExpression enumerableIndexerExpression => AutoMapper.GetOrCreateArrayIndexerMap(currentMap, enumerableIndexerExpression.Type, currentExpression.Context, enumerableIndexerExpression.ActualElementType),
            MappedMultidimensionalArrayIndexerExpression multidimensionalArrayIndexerExpression => AutoMapper.GetOrCreateMultidimensionalIndexerMap(currentMap, multidimensionalArrayIndexerExpression.Type, currentExpression.Context, multidimensionalArrayIndexerExpression.ActualElementType),
            _ => throw new ArgumentException($"GetNextMap called with unsupported expression type: {nextExpression.GetType()}", nameof(nextExpression)),
        };
    }

    private static IMap GetOrCreateMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression, Func<IMappedExpression, FallbackStrategy, IMap> memberMapCreator)
    {
        var stack = BuildExpressionStack(expression);
        if (stack.Count == 0)
        {
            throw new ArgumentException($"Expression must contain a member access, indexer, or dictionary access.", nameof(expression));
        }

        return ProcessExpression<T, TProperty>(stack, classMap, stack.Pop(), memberMapCreator, classMap.EmptyValueStrategy);
    }

    public static OneToOneMap<TProperty> GetOrCreateOneToOneMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(IMappedExpression finalExpression, FallbackStrategy emptyValueStrategy)
        {
            var member = finalExpression is MappedMemberExpression memberExpression ? memberExpression.Member : null;
            return AutoMapper.CreateOneToOneMap<TProperty>(member, finalExpression.GetDefaultCellReaderFactory(), emptyValueStrategy, isAutoMapping: false)!;
        }

        return (OneToOneMap<TProperty>)GetOrCreateMap<T, TProperty>(classMap, expression, memberMapCreator);
    }

    public static ManyToOneEnumerableMap<TElement> GetOrCreateManyToOneEnumerableMap<T, TElement>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(IMappedExpression finalExpression, FallbackStrategy emptyValueStrategy)
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
            if (!AutoMapper.TryCreateSplitMapGeneric<TElement>(member, type, finalExpression.GetDefaultCellsReaderFactory() ?? new CharSplitReaderFactory(finalExpression.GetDefaultCellReaderFactory()), emptyValueStrategy, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type \"{type}\". It must be a single dimensional array, be assignable from List<T> or implement ICollection<T>.");
            }

            return map;
        }
        
        return (ManyToOneEnumerableMap<TElement>)GetOrCreateMap<T, TElement>(classMap, expression, memberMapCreator);
    }

    public static ManyToOneDictionaryMap<TKey, TValue> GetOrCreateManyToOneDictionaryMap<T, TKey, TValue>(ExcelClassMap<T> classMap, Expression expression) where TKey : notnull
    {
        static IMap memberMapCreator(IMappedExpression finalExpression, FallbackStrategy emptyValueStrategy)
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
            if (!AutoMapper.TryCreateGenericDictionaryMap<TKey, TValue>(member, type, emptyValueStrategy, isAutoMapping: false, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type \"{type}\".");
            }

            return map;
        }

        return (ManyToOneDictionaryMap<TKey, TValue>)GetOrCreateMap<T, TValue>(classMap, expression, memberMapCreator);
    }

    public static ExcelClassMap<TElement> GetOrCreateObjectMap<T, TElement>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(IMappedExpression finalExpression, FallbackStrategy emptyValueStrategy)
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
    ICellReaderFactory GetDefaultCellReaderFactory();
    ICellsReaderFactory? GetDefaultCellsReaderFactory();
}

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

    public ICellReaderFactory GetDefaultCellReaderFactory()
        => MemberMapper.GetDefaultCellReaderFactory(Member);

    public ICellsReaderFactory? GetDefaultCellsReaderFactory() => MemberMapper.GetDefaultCellsReaderFactory(Member);
}

internal class MappedEnumerableIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type ActualElementType { get; }
    public Type MappedElementType { get; }
    public int Index { get; }
    public Type MappedValueType => MappedElementType;
    public object Context => Index;

    public MappedEnumerableIndexerExpression(BinaryExpression expression, Type? mappedElementType)
    {
        Type = expression.Left.Type!;
        ActualElementType = Type.GetElementType()!;
        MappedElementType = mappedElementType ?? ActualElementType;

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
        ActualElementType = Type.GetElementTypeOrEnumerableType()!;
        MappedElementType = mappedElementType ?? ActualElementType;

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

    public ICellReaderFactory GetDefaultCellReaderFactory()
        => new ColumnIndexReaderFactory(Index);

    public ICellsReaderFactory? GetDefaultCellsReaderFactory() => null;
}

internal class MappedMultidimensionalArrayIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type ActualElementType { get; }
    public Type MappedElementType { get; }
    public int[] Indices { get; }
    public Type MappedValueType => MappedElementType;
    public object Context => Indices;

    public MappedMultidimensionalArrayIndexerExpression(MethodCallExpression expression, Type? mappedElementType)
    {
        Type = expression.Object!.Type;
        ActualElementType = Type.GetElementType()!;
        MappedElementType = mappedElementType ?? ActualElementType;

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

    public ICellReaderFactory GetDefaultCellReaderFactory()
        => new ColumnIndexReaderFactory(0);

    public ICellsReaderFactory? GetDefaultCellsReaderFactory() => null;
}

internal class MappedDictionaryIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type ValueType { get; }
    public Type ActualKeyType { get; }
    public Type ActualValueType { get; }
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
        AutoMapper.TryGetDictionaryKeyValueType(Type, out var actualKeyType, out var actualValueType);

        ActualKeyType = actualKeyType!;
        ActualValueType = actualValueType!;

        ValueType = valueType ?? ActualValueType;

        Key = constantExpression.Value!;
    }

    public ICellReaderFactory GetDefaultCellReaderFactory()
    {
        if (Key is string keyString)
        {
            if (keyString.Length == 0)
            {
                return new ColumnIndexReaderFactory(0);
            }

            return new ColumnNameReaderFactory(keyString);
        }
        else if (Key is int keyInt)
        {
            return new ColumnIndexReaderFactory(keyInt);
        }

        return new ColumnNameReaderFactory(Key.ToString()!);
    }

    public ICellsReaderFactory? GetDefaultCellsReaderFactory() => null;
}
