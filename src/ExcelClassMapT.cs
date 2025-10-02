using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Dynamic;
using System.Linq.Expressions;
using System.Reflection;
using ExcelMapper.Utilities;
using ExcelMapper.Mappers;
using System.Linq;
using ExcelMapper.Abstractions;
using System.Collections;

namespace ExcelMapper;

/// <summary>
/// A map that maps a row of a sheet to an object of the given type.
/// </summary>
/// <typeparam name="T">The typ eof the object to create.</typeparam>
public class ExcelClassMap<T> : ExcelClassMap
{
    /// <summary>
    /// Constructs the default class map for the given type.
    /// </summary>
    public ExcelClassMap() : base(typeof(T))
    {
    }

    /// <summary>
    /// Constructs a new class map for the given type using the given default strategy to use
    /// when the value of a cell is empty.
    /// </summary>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    public ExcelClassMap(FallbackStrategy emptyValueStrategy) : this()
    {
        if (!Enum.IsDefined(typeof(FallbackStrategy), emptyValueStrategy))
        {
            throw new ArgumentException($"Invalid value \"{emptyValueStrategy}\".", nameof(emptyValueStrategy));
        }

        EmptyValueStrategy = emptyValueStrategy;
    }

    /// <summary>
    /// Gets the default strategy to use when the value of a cell is empty.
    /// </summary>
    public FallbackStrategy EmptyValueStrategy { get; }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map primitives such as string, int etc.
    /// </summary>
    /// <typeparam name="TProperty">The type of the property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public OneToOneMap<TProperty> Map<TProperty>(Expression<Func<T, TProperty>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        var map = AutoMapper.CreateMemberMap<TProperty>(memberExpression.Member, EmptyValueStrategy, false)!;
        AddMap(new ExcelPropertyMap<T>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map enums.
    /// </summary>
    /// <typeparam name="TProperty">The enum type of the property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <param name="ignoreCase">A flag indicating whether enum parsing is case insensitive.</param>
    /// <returns>The map for the given property or field.</returns>
    public OneToOneMap<TProperty> Map<TProperty>(Expression<Func<T, TProperty>> expression, bool ignoreCase) where TProperty : struct
    {
        if (!typeof(TProperty).GetTypeInfo().IsEnum)
        {
            throw new ArgumentException($"The type ${typeof(TProperty)} must be an Enum.", nameof(TProperty));
        }

        var mapper = new EnumMapper(typeof(TProperty), ignoreCase);
        var map = Map(expression);
        map.RemoveCellValueMapper(0);
        map.AddCellValueMapper(mapper);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map nullable enums.
    /// </summary>
    /// <typeparam name="TProperty">The enum type of the property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <param name="ignoreCase">A flag indicating whether enum parsing is case insensitive.</param>
    /// <returns>The map for the given property or field.</returns>
    public OneToOneMap<TProperty?> Map<TProperty>(Expression<Func<T, TProperty?>> expression, bool ignoreCase) where TProperty : struct
    {
        if (!typeof(TProperty).GetTypeInfo().IsEnum)
        {
            throw new ArgumentException($"The type ${typeof(TProperty)} must be an Enum.", nameof(TProperty));
        }

        var mapper = new EnumMapper(typeof(TProperty), ignoreCase);
        var map = Map(expression);
        map.RemoveCellValueMapper(0);
        map.AddCellValueMapper(mapper);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map enumerables.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> MapList<TElement>(Expression<Func<T, IEnumerable>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map enumerables.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, IEnumerable<TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map ICollections.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, ICollection<TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map Collection.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, Collection<TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map ObservableCollection.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, ObservableCollection<TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map ILists.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, IList<TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IReadOnlyCollections.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlyCollection<TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IReadOnlyLists.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlyList<TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map lists.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, List<TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map arrays.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, TElement[]>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        ManyToOneEnumerableMap<TElement> map = GetMultiMap<TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map objects that contain nested objects, primitives or enumerables.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ExcelClassMap<TElement> MapObject<TElement>(Expression<Func<T, TElement>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        if (!AutoMapper.TryCreateObjectMap(EmptyValueStrategy, out ExcelClassMap<TElement>? map))
        {
            throw new ExcelMappingException($"Could not map object of type \"{typeof(TElement)}\".");
        }

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IDictionarys.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<TElement> Map<TElement>(Expression<Func<T, IDictionary<string, TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        var map = GetDictionaryMap<string, TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IReadOnlyDictionarys.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlyDictionary<string, TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        var map = GetDictionaryMap<string, TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IDictionarys.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<TElement> Map<TElement>(Expression<Func<T, Dictionary<string, TElement>>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        var map = GetDictionaryMap<string, TElement>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<TElement>(memberExpression.Member, map), expression);
        return map;
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map ExpandoObjects.
    /// </summary>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<object> Map(Expression<Func<T, ExpandoObject>> expression)
    {
        MemberExpression memberExpression = GetMemberExpression(expression);
        var map = GetDictionaryMap<string, object>(memberExpression.Member);

        AddMap(new ExcelPropertyMap<object>(memberExpression.Member, map), expression);
        return map;
    }

    private ManyToOneEnumerableMap<TElement> GetMultiMap<TElement>(MemberInfo member)
    {
        if (!AutoMapper.TryCreateSplitMap<TElement>(member, EmptyValueStrategy, out var map))
        {
            throw new ExcelMappingException($"No known way to instantiate type \"{member.MemberType()}\". It must be a single dimensional array, be assignable from List<T> or implement ICollection<T>.");
        }

        return (ManyToOneEnumerableMap<TElement>)map;
    }

    private ManyToOneDictionaryMap<TValue> GetDictionaryMap<TKey, TValue>(MemberInfo member)
    {
        if (!AutoMapper.TryCreateGenericDictionaryMap<TKey, TValue>(member.MemberType(), EmptyValueStrategy, out ManyToOneDictionaryMap<TValue>? map))
        {
            throw new ExcelMappingException($"No known way to instantiate type \"{member.MemberType()}\".");
        }

        return map;
    }

    private Expression GetExpressionBody(Expression body)
    {
        // Allow casts e.g. (p => (ICollection<int>)p.Value)
        if (body is UnaryExpression unaryExpression && unaryExpression.NodeType == ExpressionType.Convert)
        {
            return unaryExpression.Operand;
        }

        return body;
    }

    protected internal MemberExpression GetMemberExpression<TElement>(Expression<Func<T, TElement>> expression)
    {
        Expression body = GetExpressionBody(expression.Body);
        if (body is not MemberExpression rootMemberExpression)
        {
            throw new ArgumentException($"Not a member expression. Received {expression.Body}", nameof(expression));
        }

        return rootMemberExpression;
    }

    protected internal void AddMap<TElement>(ExcelPropertyMap map, Expression<Func<T, TElement>> expression)
    {
        Expression expressionBody = GetExpressionBody(expression.Body);
        var expressions = new Stack<MemberExpression>();
        while (true)
        {
            if (expressionBody is not MemberExpression memberExpressionBody)
            {
                // Each map is of the form (parameter => member).
                if (expressionBody is ParameterExpression _)
                {
                    break;
                }

                throw new ArgumentException($"Expression can only contain member accesses, but found {expressionBody}. Found {expressions.Count} expressions.", nameof(expression));
            }

            expressions.Push(memberExpressionBody);
            expressionBody = memberExpressionBody.Expression;
        }

        if (expressions.Count == 1)
        {
            // Simple case: parameter => prop
            Properties.Add(map);
        }
        else
        {
            // Go through the chain of members and make sure that they are valid.
            // E.g. parameter => parameter.prop.subprop.field.
            CreateObjectMap(map, expressions);
        }
    }

    /// <summary>
    /// Traverses through a list of member expressions, starting with the member closest to the type
    /// of this class map, and creates a map for each sub member access.
    /// This enables support for expressions such as p => p.prop.subprop.field.final.
    /// </summary>
    /// <param name="propertyMap">The map for the final member access in the stack.</param>
    /// <param name="memberExpressions">A stack of each MemberExpression in the list of member access expressions.</param>
    protected internal void CreateObjectMap(ExcelPropertyMap propertyMap, Stack<MemberExpression> memberExpressions)
    {
        MemberExpression memberExpression = memberExpressions.Pop();
        if (memberExpressions.Count == 0)
        {
            // This is the final member.
            Properties.Add(propertyMap);
            return;
        }

        Type memberType = memberExpression.Member.MemberType();

        MethodInfo method = MapObjectMethod.MakeGenericMethod(memberType);
        try
        {
            method.Invoke(this, [propertyMap, memberExpression, memberExpressions]);
        }
        catch (TargetInvocationException exception)
        {
            // Discarding InnerException's nullability warning
            // because it will never be null.
            // It is nullable only because base Exception has it as nullable.
            throw exception.InnerException!;
        }
    }

    private void CreateObjectMapGeneric<TElement>(ExcelPropertyMap propertyMap, MemberExpression memberExpression, Stack<MemberExpression> memberExpressions)
    {
        ExcelPropertyMap? map = Properties.FirstOrDefault(m => m.Member.Equals(memberExpression.Member));

        ExcelClassMap<TElement> objectPropertyMap;
        if (map == null)
        {
            objectPropertyMap = new ExcelClassMap<TElement>();
            Properties.Add(new ExcelPropertyMap(memberExpression.Member, objectPropertyMap));
        }
        else if (map.Map is not ExcelClassMap<TElement> existingMap)
        {
            throw new InvalidOperationException($"Expression is already mapped differently as {map.GetType()}.");
        }
        else
        {
            objectPropertyMap = existingMap;
        }

        objectPropertyMap.CreateObjectMap(propertyMap, memberExpressions);
    }

    /// <summary>
    /// Configures the existing class map used to map multiple cells in a row to the properties and fields
    /// of a an object.
    /// </summary>
    /// <param name="classMapFactory">A delegate that allows configuring the default class map used.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ExcelClassMap<T> WithClassMap(Action<ExcelClassMap<T>> classMapFactory)
    {
        if (classMapFactory == null)
        {
            throw new ArgumentNullException(nameof(classMapFactory));
        }

        classMapFactory(this);
        return this;
    }

    /// <summary>
    /// Sets the new class map used to map multiple cells in a row to the properties and fields
    /// of a an object.
    /// </summary>
    /// <param name="classMap">The new class map used.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ExcelClassMap<T> WithClassMap(ExcelClassMap<T> classMap)
    {
        if (classMap == null)
        {
            throw new ArgumentNullException(nameof(classMap));
        }

        Properties.Clear();
        foreach (ExcelPropertyMap propertyMap in classMap.Properties)
        {
            Properties.Add(propertyMap);
        }

        return this;
    }

    private static MethodInfo? s_mapObjectMethod;
    private static MethodInfo MapObjectMethod => s_mapObjectMethod ?? (s_mapObjectMethod = typeof(ExcelClassMap<T>).GetTypeInfo().GetDeclaredMethod(nameof(CreateObjectMapGeneric)));
}
