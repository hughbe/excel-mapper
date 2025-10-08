using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Dynamic;
using System.Linq.Expressions;
using ExcelMapper.Utilities;
using ExcelMapper.Mappers;
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
    public ExcelClassMap(FallbackStrategy emptyValueStrategy) : base(typeof(T), emptyValueStrategy)
    {
    }

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map primitives such as string, int etc.
    /// </summary>
    /// <typeparam name="TProperty">The type of the property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public OneToOneMap<TProperty> Map<TProperty>(Expression<Func<T, TProperty>> expression)
        => GetOrCreateOneToOneMap<TProperty>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map enums.
    /// </summary>
    /// <typeparam name="TProperty">The enum type of the property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <param name="ignoreCase">A flag indicating whether enum parsing is case insensitive.</param>
    /// <returns>The map for the given property or field.</returns>
    public OneToOneMap<TProperty> Map<TProperty>(Expression<Func<T, TProperty>> expression, bool ignoreCase) where TProperty : struct
        => MapEnumInternal<TProperty>(expression, ignoreCase);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map nullable enums.
    /// </summary>
    /// <typeparam name="TProperty">The enum type of the property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <param name="ignoreCase">A flag indicating whether enum parsing is case insensitive.</param>
    /// <returns>The map for the given property or field.</returns>
    public OneToOneMap<TProperty?> Map<TProperty>(Expression<Func<T, TProperty?>> expression, bool ignoreCase) where TProperty : struct
        => MapEnumInternal<TProperty?>(expression, ignoreCase);

    private OneToOneMap<TProperty> MapEnumInternal<TProperty>(Expression<Func<T, TProperty>> expression, bool ignoreCase)
    {
        var enumType = typeof(TProperty).GetNullableTypeOrThis(out _);
        if (!enumType.IsEnum)
        {
            throw new ArgumentException($"The type ${enumType} must be an Enum.", nameof(TProperty));
        }

        var mapper = new EnumMapper(enumType, ignoreCase);
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
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map enumerables.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, IEnumerable<TElement>>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map ICollections.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, ICollection<TElement>>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map Collection.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, Collection<TElement>>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map ObservableCollection.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, ObservableCollection<TElement>>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map ILists.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, IList<TElement>>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IReadOnlyCollections.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlyCollection<TElement>>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IReadOnlyLists.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlyList<TElement>>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map lists.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, List<TElement>>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map arrays.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneEnumerableMap<TElement> Map<TElement>(Expression<Func<T, TElement[]>> expression)
        => GetOrCreateManyToOneEnumerableMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map objects that contain nested objects, primitives or enumerables.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ExcelClassMap<TElement> MapObject<TElement>(Expression<Func<T, TElement>> expression)
        => GetOrCreateObjectMap<TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IDictionarys.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<TElement> Map<TElement>(Expression<Func<T, IDictionary>> expression)
        => GetOrCreateManyToOneDictionaryMap<string, TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IDictionarys.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<TElement> Map<TElement>(Expression<Func<T, IDictionary<string, TElement>>> expression)
        => GetOrCreateManyToOneDictionaryMap<string, TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IReadOnlyDictionarys.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlyDictionary<string, TElement>>> expression)
        => GetOrCreateManyToOneDictionaryMap<string, TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map IDictionarys.
    /// </summary>
    /// <typeparam name="TElement">The element type of property or field to map.</typeparam>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<TElement> Map<TElement>(Expression<Func<T, Dictionary<string, TElement>>> expression)
        => GetOrCreateManyToOneDictionaryMap<string, TElement>(expression.Body);

    /// <summary>
    /// Creates a map for a property or field given a MemberExpression reading the property or field.
    /// This is used for map ExpandoObjects.
    /// </summary>
    /// <param name="expression">A MemberExpression reading the property or field.</param>
    /// <returns>The map for the given property or field.</returns>
    public ManyToOneDictionaryMap<object> Map(Expression<Func<T, ExpandoObject>> expression)
        => GetOrCreateManyToOneDictionaryMap<string, object>(expression.Body);

    // Mapping methods now use ExpressionAutoMapper static methods
    private OneToOneMap<TProperty> GetOrCreateOneToOneMap<TProperty>(Expression expression)
        => ExpressionAutoMapper.GetOrCreateOneToOneMap<T, TProperty>(this, expression);

    private ManyToOneEnumerableMap<TElement> GetOrCreateManyToOneEnumerableMap<TElement>(Expression expression)
        => ExpressionAutoMapper.GetOrCreateManyToOneEnumerableMap<T, TElement>(this, expression);

    private ExcelClassMap<TElement> GetOrCreateObjectMap<TElement>(Expression expression)
        => ExpressionAutoMapper.GetOrCreateObjectMap<T, TElement>(this, expression);

    private ManyToOneDictionaryMap<TValue> GetOrCreateManyToOneDictionaryMap<TKey, TValue>(Expression expression) where TKey : notnull
        => ExpressionAutoMapper.GetOrCreateManyToOneDictionaryMap<T, TKey, TValue>(this, expression);

    /// <summary>
    /// Configures the existing class map used to map multiple cells in a row to the properties and fields
    /// of a an object.
    /// </summary>
    /// <param name="classMapFactory">A delegate that allows configuring the default class map used.</param>
    /// <returns>The map that invoked this method.</returns>
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
    /// <returns>The map that invoked this method.</returns>
    public ExcelClassMap<T> WithClassMap(ExcelClassMap<T> classMap)
    {
        if (classMap == null)
        {
            throw new ArgumentNullException(nameof(classMap));
        }

        Properties.Clear();
        foreach (var propertyMap in classMap.Properties)
        {
            Properties.Add(propertyMap);
        }

        return this;
    }
}
