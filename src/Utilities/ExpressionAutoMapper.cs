using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Linq;
using System.Linq.Expressions;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Utilities;

internal static class ExpressionAutoMapper
{
    private static Expression SkipConvert(Expression expression, out Type convertedType)
    {
        convertedType = null!;
        while (expression is UnaryExpression unaryExpression && unaryExpression.NodeType == ExpressionType.Convert)
        {
            expression = unaryExpression.Operand;
            // If we have multiple converts, we want the innermost type
            // e.g. ((A)(B)x) should give us A, not B
            if (convertedType is null)
            {
                convertedType = unaryExpression.Type;
            }
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
                expressions.Push(new MappedMemberExpression(memberExpression, currentTargetType));
                currentExpression = memberExpression.Expression!;
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
            methodCallExpression.Arguments.All(arg => arg is ConstantExpression constantExpr && constantExpr.Type == typeof(int));

    private static bool IsListIndexerExpression(Expression expression)
        => expression is MethodCallExpression methodCallExpression &&
            methodCallExpression.Method.Name == "get_Item" &&
            methodCallExpression.Object is not null &&
            methodCallExpression.Object!.Type!.GetElementTypeOrEnumerableType() != null &&
            methodCallExpression.Arguments.Count == 1 &&
            methodCallExpression.Arguments[0] is ConstantExpression &&
            methodCallExpression.Arguments[0].Type == typeof(int);

    private static bool IsDictionaryIndexerExpression(Expression expression)
        => expression is MethodCallExpression methodCallExpression &&
            methodCallExpression.Object is not null &&
            AutoMapper.TryGetDictionaryKeyValueType(methodCallExpression.Object!.Type!, out var _, out var _) &&
            methodCallExpression.Method.Name == "get_Item" &&
            methodCallExpression.Arguments.Count == 1 &&
            methodCallExpression.Arguments[0] is ConstantExpression;

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

    private static IMap CreateAndAddElementMap<TProperty>(ExcelClassMap classMap, IMap currentMap, IMappedExpression expression, int index, Type valueType)
    {
        var map = AutoMapper.CreateArrayIndexerElementMap(index, valueType, classMap.EmptyValueStrategy);
        AddMap(currentMap, expression, map);
        return map;
    }

    private static IMap CreateAndAddElementMap<TProperty>(ExcelClassMap classMap, IMap currentMap, IMappedExpression expression, int[] indices, Type valueType)
    {
        var map = AutoMapper.CreateArrayIndexerElementMap(0, valueType, classMap.EmptyValueStrategy);
        AddMap(currentMap, expression, map);
        return map;
    }

    private static IMap CreateAndAddElementMap<TProperty>(ExcelClassMap classMap, IMap currentMap, IMappedExpression expression, object key, Type valueType)
    {
        var map = AutoMapper.CreateDictionaryIndexerElementMap(key, valueType, classMap.EmptyValueStrategy);
        AddMap(currentMap, expression, map);
        return map;
    }

    [ExcludeFromCodeCoverage]
    private static IMap ProcessExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<IMappedExpression> stack, IMap currentMap, IMappedExpression currentExpression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        return currentExpression switch
        {
            MappedMemberExpression memberExpression => ProcessMemberExpression<T, TProperty>(classMap, stack, currentMap, memberExpression, memberMapCreator),
            MappedDictionaryIndexerExpression dictionaryIndexerExpression => ProcessDictionaryIndexerExpression<T, TProperty>(classMap, stack, currentMap, dictionaryIndexerExpression, memberMapCreator),
            MappedEnumerableIndexerExpression enumerableIndexerExpression => ProcessEnumerableIndexExpression<T, TProperty>(classMap, stack, currentMap, enumerableIndexerExpression, memberMapCreator),
            MappedMultidimensionalArrayIndexerExpression multidimensionalArrayIndexerExpression => ProcessMultidimensionalArrayIndexerExpression<T, TProperty>(classMap, stack, currentMap, multidimensionalArrayIndexerExpression, memberMapCreator),
            _ => throw new ArgumentException($"Unsupported expression type: {currentExpression.GetType()}", nameof(currentExpression)),
        };
    }

    private static IMap ProcessMemberExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<IMappedExpression> stack, IMap currentMap, MappedMemberExpression memberExpression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        if (stack.Count == 0)
        {
            var memberMap = memberMapCreator(memberExpression.Member, ((ExcelClassMap)currentMap).EmptyValueStrategy);
            AddMap(currentMap, memberExpression, memberMap);
            return memberMap;
        }

        var nextExpression = stack.Peek();
        var nextMap = GetNextMap(classMap, currentMap, memberExpression.Type, nextExpression, memberExpression.Member);
        return ProcessNextExpression<T, TProperty>(classMap, stack, nextMap, memberMapCreator);
    }

    private static IMap ProcessEnumerableIndexExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<IMappedExpression> stack, IMap currentMap, MappedEnumerableIndexerExpression expression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        if (stack.Count == 0)
        {
            return CreateAndAddElementMap<TProperty>(classMap, currentMap, expression, expression.Index, expression.MappedElementType);
        }

        var nextExpression = stack.Peek();
        var nextMap = GetNextMap(classMap, currentMap, expression.MappedElementType, nextExpression, null, expression.Index, null);
        return ProcessNextExpression<T, TProperty>(classMap, stack, nextMap, memberMapCreator);
    }

    private static IMap ProcessMultidimensionalArrayIndexerExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<IMappedExpression> stack, IMap currentMap, MappedMultidimensionalArrayIndexerExpression expression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        if (stack.Count == 0)
        {
            return CreateAndAddElementMap<TProperty>(classMap, currentMap, expression, expression.Indices, expression.MappedElementType);
        }

        var nextExpression = stack.Peek();
        var nextMap = GetNextMap(classMap, currentMap, expression.MappedElementType, nextExpression, null, null, expression.Indices);
        return ProcessNextExpression<T, TProperty>(classMap, stack, nextMap, memberMapCreator);
    }

    private static IMap ProcessDictionaryIndexerExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<IMappedExpression> stack, IMap currentMap, MappedDictionaryIndexerExpression expression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        if (stack.Count == 0)
        {
            return CreateAndAddElementMap<TProperty>(classMap, currentMap, expression, expression.Key, expression.ValueType);
        }
        var nextExpression = stack.Peek();
        var nextMap = GetNextMap(classMap, currentMap, expression.ValueType, nextExpression, null, null, expression.Key);
        return ProcessNextExpression<T, TProperty>(classMap, stack, nextMap, memberMapCreator);
    }

    private static IMap ProcessNextExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<IMappedExpression> stack, IMap currentMap, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        var nextExpression = stack.Pop();
        return ProcessExpression<T, TProperty>(classMap, stack, currentMap, nextExpression, memberMapCreator);
    }

    [ExcludeFromCodeCoverage]
    private static IMap GetNextMap(ExcelClassMap classMap, IMap currentMap, Type currentValueType, IMappedExpression nextExpression, MemberInfo? member = null, int? index = null, object? key = null)
    {
        return nextExpression switch
        {
            MappedMemberExpression memberExpression => AutoMapper.GetOrCreateNestedMap(currentMap, member, currentValueType, index ?? key),
            MappedDictionaryIndexerExpression dictionaryIndexerExpression => AutoMapper.GetOrCreateDictionaryIndexerMap(currentMap, member, dictionaryIndexerExpression.Type, index ?? key, dictionaryIndexerExpression.ActualKeyType, dictionaryIndexerExpression.ActualValueType),
            MappedEnumerableIndexerExpression enumerableIndexerExpression => AutoMapper.GetOrCreateArrayIndexerMap(currentMap, member, enumerableIndexerExpression.Type, index ?? key, enumerableIndexerExpression.ActualElementType),
            MappedMultidimensionalArrayIndexerExpression multidimensionalArrayIndexerExpression => AutoMapper.GetOrCreateMultidimensionalIndexerMap(currentMap, member, multidimensionalArrayIndexerExpression.Type, index ?? key, multidimensionalArrayIndexerExpression.ActualElementType),
            _ => throw new ArgumentException($"GetNextMap called with unsupported expression type: {nextExpression.GetType()}", nameof(nextExpression)),
        };
    }

    private static IMap GetOrCreateMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        var stack = BuildExpressionStack(expression);
        IMap currentMap = classMap;
        if (stack.TryPop(out var currentExpression))
        {
            return ProcessExpression<T, TProperty>(classMap, stack, currentMap, currentExpression, memberMapCreator);
        }

        throw new ArgumentException($"Expression must contain a member access, indexer, or dictionary access. Received: {expression}", nameof(expression));
    }

    public static OneToOneMap<TProperty> GetOrCreateOneToOneMap<T, TProperty>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(MemberInfo member, FallbackStrategy emptyValueStrategy)
        {
            return AutoMapper.CreateMemberMap<TProperty>(member, emptyValueStrategy, isAutoMapping: false)!;
        }
        return (OneToOneMap<TProperty>)GetOrCreateMap<T, TProperty>(classMap, expression, memberMapCreator);
    }

    public static ManyToOneEnumerableMap<TElement> GetOrCreateManyToOneEnumerableMap<T, TElement>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(MemberInfo member, FallbackStrategy emptyValueStrategy)
        {
            if (!AutoMapper.TryCreateSplitMap<TElement>(member, emptyValueStrategy, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type \"{member.MemberType()}\". It must be a single dimensional array, be assignable from List<T> or implement ICollection<T>.");
            }

            return map;
        }
        return (ManyToOneEnumerableMap<TElement>)GetOrCreateMap<T, TElement>(classMap, expression, memberMapCreator);
    }

    public static ManyToOneDictionaryMap<TKey, TValue> GetOrCreateManyToOneDictionaryMap<T, TKey, TValue>(ExcelClassMap<T> classMap, Expression expression) where TKey : notnull
    {
        static IMap memberMapCreator(MemberInfo member, FallbackStrategy emptyValueStrategy)
        {
            if (!AutoMapper.TryCreateGenericDictionaryMap<TKey, TValue>(member, member.MemberType(), emptyValueStrategy, isAutoMapping: false, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type \"{member.MemberType()}\".");
            }

            return map;
        }
        return (ManyToOneDictionaryMap<TKey, TValue>)GetOrCreateMap<T, TValue>(classMap, expression, memberMapCreator);
    }

    public static ExcelClassMap<TElement> GetOrCreateObjectMap<T, TElement>(ExcelClassMap<T> classMap, Expression expression)
    {
        static IMap memberMapCreator(MemberInfo member, FallbackStrategy emptyValueStrategy)
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
}

internal class MappedMemberExpression : IMappedExpression
{
    public Type Type { get; set; }
    public MemberInfo Member { get; }

    public MappedMemberExpression(MemberExpression expression, Type memberType)
    {
        Member = expression.Member;
        Type = memberType ?? expression.Type;
    }
}

internal class MappedEnumerableIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type ActualElementType { get; }
    public Type MappedElementType { get; }
    public int Index { get; }

    public MappedEnumerableIndexerExpression(BinaryExpression expression, Type? mappedElementType)
    {
        Type = expression.Left.Type!;
        ActualElementType = Type.GetElementType()!;
        MappedElementType = mappedElementType ?? ActualElementType;

        if (expression.Right is not ConstantExpression constantExpression)
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
        Type = expression.Object!.Type;
        ActualElementType = Type.GetElementTypeOrEnumerableType()!;
        MappedElementType = mappedElementType ?? ActualElementType;

        var constantExpression = (ConstantExpression)expression.Arguments[0];
        if ((int)constantExpression.Value! < 0)
        {
            throw new ArgumentException($"Array and list indexers must be non-negative. Received: {constantExpression.Value!}", nameof(expression));
        }

        Index = (int)constantExpression.Value!;
    }
}

internal class MappedMultidimensionalArrayIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type ActualElementType { get; }
    public Type MappedElementType { get; }
    public int[] Indices { get; }

    public MappedMultidimensionalArrayIndexerExpression(MethodCallExpression expression, Type? mappedElementType)
    {
        Type = expression.Object!.Type;
        ActualElementType = Type.GetElementType()!;
        MappedElementType = mappedElementType ?? ActualElementType;

        var arguments = expression.Arguments;
        Indices = new int[arguments.Count];
        for (int i = 0; i < arguments.Count; i++)
        {
            var constantExpression = (ConstantExpression)arguments[i];
            if ((int)constantExpression.Value! < 0)
            {
                throw new ArgumentException($"Array indices must be non-negative. Received: {constantExpression.Value!}", nameof(expression));
            }

            Indices[i] = (int)constantExpression.Value!;
        }
    }
}

internal class MappedDictionaryIndexerExpression : IMappedExpression
{
    public Type Type { get; }
    public Type ValueType { get; }
    public Type ActualKeyType { get; }
    public Type ActualValueType { get; }
    public object Key { get; }

    public MappedDictionaryIndexerExpression(MethodCallExpression expression, Type valueType)
    {
        Type = expression.Object!.Type;
        AutoMapper.TryGetDictionaryKeyValueType(Type, out var actualKeyType, out var actualValueType);

        ActualKeyType = actualKeyType!;
        ActualValueType = actualValueType!;

        ValueType = valueType ?? ActualValueType;

        var constantExpression = (ConstantExpression)expression.Arguments[0];
        if (constantExpression.Value == null)
        {
            throw new ArgumentException($"Dictionary indexer key cannot be null. Received: {constantExpression.Value!}", nameof(expression));
        }

        Key = constantExpression.Value!;
    }
}
