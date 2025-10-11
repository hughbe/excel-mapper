using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Linq.Expressions;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Utilities;

internal static class ExpressionAutoMapper
{
    private static Stack<Expression> BuildExpressionStack(Expression expression)
    {
        var expressions = new Stack<Expression>();
        var currentExpression = expression;
        while (true)
        {
            if (currentExpression is MemberExpression memberExpressionBody)
            {
                expressions.Push(currentExpression);
                currentExpression = memberExpressionBody.Expression!;
            }
            else if (currentExpression is UnaryExpression unaryExpression && unaryExpression.NodeType == ExpressionType.Convert)
            {
                currentExpression = unaryExpression.Operand;
            }
            else if (IsArrayIndexerExpression(currentExpression))
            {
                var binaryExpression = (BinaryExpression)currentExpression;
                expressions.Push(currentExpression);
                currentExpression = binaryExpression.Left;
            }
            else if (IsListIndexerExpression(currentExpression))
            {
                var methodCallExpression = (MethodCallExpression)currentExpression;
                expressions.Push(currentExpression);
                currentExpression = methodCallExpression.Object!;
            }
            else if (IsDictionaryIndexerExpression(currentExpression))
            {
                var methodCallExpression = (MethodCallExpression)currentExpression;
                expressions.Push(currentExpression);
                currentExpression = methodCallExpression.Object!;
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

    private static bool IsListIndexerExpression(Expression expression)
        => expression is MethodCallExpression methodCallExpression && methodCallExpression.Method.Name == "get_Item" && methodCallExpression.Arguments.Count == 1 && methodCallExpression.Arguments[0] is ConstantExpression && methodCallExpression.Arguments[0].Type == typeof(int);

    private static bool IsDictionaryIndexerExpression(Expression expression)
        => expression is MethodCallExpression methodCallExpression && methodCallExpression.Method.Name == "get_Item" && methodCallExpression.Arguments.Count == 1 && methodCallExpression.Arguments[0] is ConstantExpression && methodCallExpression.Arguments[0].Type == typeof(string);

    private static int GetIndexFromArrayOrListIndexer(Expression expression)
    {
        Expression indexExpression;
        if (IsArrayIndexerExpression(expression))
        {
            var binaryExpression = (BinaryExpression)expression;
            indexExpression = binaryExpression.Right;
        }
        else // IsListIndexerExpression
        {
            var methodCallExpression = (MethodCallExpression)expression;
            indexExpression = methodCallExpression.Arguments[0];
        }

        if (indexExpression is not ConstantExpression constantExpression)
        {
            throw new ArgumentException($"The indexer must be a constant expression. Received {indexExpression}.", nameof(expression));
        }
        if ((int)constantExpression.Value! < 0)
        {
            throw new ArgumentException($"Array and list indexers must be non-negative. Received: {constantExpression.Value!}", nameof(expression));
        }

        return (int)constantExpression.Value!;
    }

    private static (int index, Type valueType) ParseArrayOrListIndexer(Expression expression)
    {
        var index = GetIndexFromArrayOrListIndexer(expression);

        Type valueType;
        if (IsArrayIndexerExpression(expression))
        {
            var binaryExpression = (BinaryExpression)expression;
            valueType = binaryExpression.Left.Type.GetElementType()!;
        }
        else // IsListIndexerExpression
        {
            var methodCallExpression = (MethodCallExpression)expression;
            valueType = methodCallExpression.Method.ReturnType;
        }

        return (index, valueType);
    }

    private static (string key, Type valueType) ParseDictionaryIndexer(Expression expression)
    {
        var methodCallExpression = (MethodCallExpression)expression;
        var constantExpression = (ConstantExpression)methodCallExpression.Arguments[0];
        var key = (string)constantExpression.Value!;
        if (key is null)
        {
            throw new ArgumentException($"Dictionary indexer key cannot be null. Received: {constantExpression.Value!}", nameof(expression));
        }

        var valueType = methodCallExpression.Method.ReturnType;
        return (key, valueType);
    }

    private static void AddMap(IMap parentMap, Expression expression, IMap map)
    {
        if (parentMap is ExcelClassMap classMap)
        {
            var member = ((MemberExpression)expression).Member;
            var existingPropertyMap = classMap.Properties.FirstOrDefault(m => m.Member.Equals(member));
            if (existingPropertyMap is not null)
            {
                classMap.Properties.Remove(existingPropertyMap);
            }

            classMap.Properties.Add(new ExcelPropertyMap(member, map));
        }
        else if (parentMap is IEnumerableIndexerMap arrayIndexerMap)
        {
            var (index, _) = ParseArrayOrListIndexer(expression);
            arrayIndexerMap.Values[index] = map;
        }
        else
        {
            var dictionaryIndexerMap = (IDictionaryIndexerMap)parentMap;
            var (key, _) = ParseDictionaryIndexer(expression);
            dictionaryIndexerMap.Values[key] = map;
        }
    }

    private static IMap CreateAndAddElementMap<TProperty>(ExcelClassMap classMap, IMap currentMap, Expression expression, int index, Type valueType)
    {
        var map = AutoMapper.CreateArrayIndexerElementMap(index, valueType, classMap.EmptyValueStrategy);
        AddMap(currentMap, expression, map);
        return map;
    }

    private static IMap CreateAndAddElementMap<TProperty>(ExcelClassMap classMap, IMap currentMap, Expression expression, string key, Type valueType)
    {
        var map = AutoMapper.CreateDictionaryIndexerElementMap(key, valueType, classMap.EmptyValueStrategy);
        AddMap(currentMap, expression, map);
        return map;
    }

    private static IMap ProcessExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<Expression> stack, IMap currentMap, Expression currentExpression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        return currentExpression switch
        {
            MemberExpression memberExpression => ProcessMemberExpression<T, TProperty>(classMap, stack, currentMap, memberExpression, memberMapCreator),
            Expression when IsArrayIndexerExpression(currentExpression) || IsListIndexerExpression(currentExpression) => ProcessArrayOrListIndexerExpression<T, TProperty>(classMap, stack, currentMap, currentExpression, memberMapCreator),
            _ => ProcessDictionaryIndexerExpression<T, TProperty>(classMap, stack, currentMap, currentExpression, memberMapCreator),
        };
    }

    private static IMap ProcessMemberExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<Expression> stack, IMap currentMap, MemberExpression memberExpression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
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

    private static IMap ProcessArrayOrListIndexerExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<Expression> stack, IMap currentMap, Expression expression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        var (index, valueType) = ParseArrayOrListIndexer(expression);
        if (stack.Count == 0)
        {
            return CreateAndAddElementMap<TProperty>(classMap, currentMap, expression, index, valueType);
        }

        var nextExpression = stack.Peek();
        var nextMap = GetNextMap(classMap, currentMap, valueType, nextExpression, null, index, null);
        return ProcessNextExpression<T, TProperty>(classMap, stack, nextMap, memberMapCreator);
    }

    private static IMap ProcessDictionaryIndexerExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<Expression> stack, IMap currentMap, Expression expression, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        var (key, valueType) = ParseDictionaryIndexer(expression);
        if (stack.Count == 0)
        {
            return CreateAndAddElementMap<TProperty>(classMap, currentMap, expression, key, valueType);
        }
        var nextExpression = stack.Peek();
        var nextMap = GetNextMap(classMap, currentMap, valueType, nextExpression, null, null, key);
        return ProcessNextExpression<T, TProperty>(classMap, stack, nextMap, memberMapCreator);
    }

    private static IMap ProcessNextExpression<T, TProperty>(ExcelClassMap<T> classMap, Stack<Expression> stack, IMap currentMap, Func<MemberInfo, FallbackStrategy, IMap> memberMapCreator)
    {
        var nextExpression = stack.Pop();
        return ProcessExpression<T, TProperty>(classMap, stack, currentMap, nextExpression, memberMapCreator);
    }

    private static IMap GetNextMap(ExcelClassMap classMap, IMap currentMap, Type currentValueType, Expression nextExpression, MemberInfo? member = null, int? index = null, object? key = null)
    {
        if (nextExpression is MemberExpression)
        {
            return AutoMapper.GetOrCreateNestedMap(currentMap, member, currentValueType, index ?? key);
        }
        else if (IsArrayIndexerExpression(nextExpression))
        {
            var nextValueType = currentValueType.GetElementType()!;
            if (member != null)
            {
                return AutoMapper.GetOrCreateArrayIndexerMap(currentMap, member, member.MemberType(), null, nextValueType);
            }
            else
            {
                return AutoMapper.GetOrCreateArrayIndexerMap(currentMap, null, currentValueType, index ?? key, nextValueType);
            }
        }
        else if (IsListIndexerExpression(nextExpression))
        {
            var nextValueType = currentValueType.GenericTypeArguments[0];
            if (member != null)
            {
                return AutoMapper.GetOrCreateArrayIndexerMap(currentMap, member, member.MemberType(), null, nextValueType);
            }
            else
            {
                return AutoMapper.GetOrCreateArrayIndexerMap(currentMap, null, currentValueType, index ?? key, nextValueType);
            }
        }
        else
        {
            var methodCallExpression = (MethodCallExpression)nextExpression;
            var keyType = methodCallExpression.Arguments[0].Type;
            var nextValueType = methodCallExpression.Method.ReturnType;
            if (member != null)
            {
                return AutoMapper.GetOrCreateDictionaryIndexerMap(currentMap, member, member.MemberType(), null, keyType, nextValueType);
            }
            else
            {
                return AutoMapper.GetOrCreateDictionaryIndexerMap(currentMap, null, currentValueType, index ?? key, keyType, nextValueType);
            }
        }
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

    public static ManyToOneDictionaryMap<TValue> GetOrCreateManyToOneDictionaryMap<T, TKey, TValue>(ExcelClassMap<T> classMap, Expression expression) where TKey : notnull
    {
        static IMap memberMapCreator(MemberInfo member, FallbackStrategy emptyValueStrategy)
        {
            if (!AutoMapper.TryCreateGenericDictionaryMap<TKey, TValue>(member, member.MemberType(), emptyValueStrategy, isAutoMapping: false, out var map))
            {
                throw new ExcelMappingException($"No known way to instantiate type \"{member.MemberType()}\".");
            }

            return map;
        }
        return (ManyToOneDictionaryMap<TValue>)GetOrCreateMap<T, TValue>(classMap, expression, memberMapCreator);
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
