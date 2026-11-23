using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics.CodeAnalysis;
using System.Dynamic;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// Maps a row of a sheet to an object of the given type.
/// </summary>
/// <typeparam name="T">The type of the object to create.</typeparam>
public class ExcelClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T> : ExcelClassMap
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
    public ExcelClassMap(EmptyValueStrategy emptyValueStrategy) : base(typeof(T), emptyValueStrategy)
    {
    }

    // Reads rows with the class map, made when the first row is read, after the class map is frozen.
    private ClassMapReader<T>? _reader;

    /// <summary>
    /// Maps a row (or for a nested object, the cells it reads) to a <typeparamref name="T"/> (see
    /// <see cref="ClassMapReader{T}.TryRead"/>).
    /// </summary>
    internal bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, out T result)
        => (_reader ??= new ClassMapReader<T>(this)).TryRead(sheet, rowIndex, reader, member, errors, out result);

    private protected sealed override bool TryGetObject(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? result)
    {
        var found = TryGetValue(sheet, rowIndex, reader, member, errors, out var value);
        result = found ? value : null;
        return found;
    }

    private protected sealed override BoundMember<TOwner> Bind<TOwner>(ClassMapMember member) => new BoundClassMember<TOwner, T>(this, member);

    private protected sealed override MethodInfo? GetArgumentMethod<TOwner>(Type parameterType)
        => parameterType == typeof(T) ? new Func<BoundMember<TOwner, T>, T, T>(BoundMember<TOwner, T>.GetValueOrDefault).Method : null;

    /// <summary>
    /// Maps a property or field read from a single cell, such as a string, number, date or enum.
    /// </summary>
    /// <typeparam name="TProperty">The type of the property or field.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Name</c>.</param>
    /// <returns>The map of the property or field, to choose its column and configure how cells are converted.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public ValueMap<TProperty> Map<TProperty>(Expression<Func<T, TProperty>> expression)
        => ExpressionAutoMapper.GetOrCreateValueMap<T, TProperty>(this, expression.Body);

    /// <summary>
    /// Maps a collection property or field of a type without its own <c>Map</c> overload, such as
    /// <see cref="HashSet{T}"/>: its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> MapList<TElement>(Expression<Func<T, IEnumerable>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="IEnumerable{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, IEnumerable<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="ICollection{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, ICollection<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps a <see cref="Collection{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, Collection<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps a <see cref="ReadOnlyCollection{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, ReadOnlyCollection<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="ObservableCollection{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, ObservableCollection<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps a <see cref="ReadOnlyObservableCollection{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, ReadOnlyObservableCollection<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="ISet{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, ISet<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

#if NET5_0_OR_GREATER
    /// <summary>
    /// Maps an <see cref="IReadOnlySet{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlySet<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);
#endif

#if NET9_0_OR_GREATER
    /// <summary>
    /// Maps a <see cref="ReadOnlySet{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, ReadOnlySet<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);
#endif

    /// <summary>
    /// Maps an <see cref="IList{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, IList<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="IReadOnlyCollection{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlyCollection<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="IReadOnlyList{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, IReadOnlyList<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps a <see cref="List{TElement}"/> property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, List<TElement>>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps an array property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, Array>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps an array property or field:
    /// its elements are read from several columns, or split from one cell.
    /// </summary>
    /// <typeparam name="TElement">The type of the elements.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Tags</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its elements.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public CollectionMap<TElement> Map<TElement>(Expression<Func<T, TElement[]>> expression)
        => ExpressionAutoMapper.GetOrCreateCollectionMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Maps a property or field that is an object with members of its own, read from the same row.
    /// </summary>
    /// <typeparam name="TElement">The type of the object.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Address</c>.</param>
    /// <returns>The class map of the object, to map its members.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public ExcelClassMap<TElement> MapObject<TElement>(Expression<Func<T, TElement>> expression)
        => ExpressionAutoMapper.GetOrCreateObjectMap<T, TElement>(this, expression.Body);

    /// <summary>
    /// Sets a property or field to the number of the row being read instead of reading a column.
    /// Row numbers are one-based, as Excel shows them, and count every row in the sheet including
    /// the heading, so they match the rows in error messages.
    /// </summary>
    /// <typeparam name="TProperty">The type of the property or field: a numeric type such as
    /// <see cref="int"/> or <see cref="long"/>, or a nullable numeric type.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Row</c>.</param>
    /// <exception cref="ArgumentException">The property or field is not a numeric type.</exception>
    /// <remarks>A row number has nothing to configure, so nothing is returned.</remarks>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public void MapRowNumber<TProperty>(Expression<Func<T, TProperty>> expression)
        => ExpressionAutoMapper.GetOrCreateRowNumberMap<T, TProperty>(this, expression.Body);

    /// <summary>
    /// Fills a dictionary property or field with every column in the heading that no other member
    /// of the same class reads, keyed by column name.
    /// </summary>
    /// <typeparam name="TValue">The type of the dictionary's values.</typeparam>
    /// <param name="expression">An expression reading a dictionary property or field with
    /// string keys, such as <see cref="Dictionary{TKey, TValue}"/> or <see cref="System.Dynamic.ExpandoObject"/>.</param>
    /// <returns>The map of the property or field, to configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, TValue> MapExtraColumns<TValue>(Expression<Func<T, IEnumerable<KeyValuePair<string, TValue>>>> expression)
        => ExpressionAutoMapper.GetOrCreateExtraColumnsMap<T, TValue>(this, expression.Body);

    /// <summary>
    /// Maps a dictionary property or field of a type without its own <c>Map</c> overload: each value is
    /// read from a column, keyed by the column's name.
    /// </summary>
    /// <typeparam name="TValue">The type of the values.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Scores</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, TValue> MapDictionary<TValue>(Expression<Func<T, IEnumerable>> expression)
        => ExpressionAutoMapper.GetOrCreateDictionaryMap<T, string, TValue>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="System.Collections.IDictionary"/> property or field:
    /// each value is read from a column, keyed by the column's name.
    /// </summary>
    /// <typeparam name="TValue">The type of the values.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Scores</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, TValue> Map<TValue>(Expression<Func<T, IDictionary>> expression)
        => ExpressionAutoMapper.GetOrCreateDictionaryMap<T, string, TValue>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="IDictionary{TKey, TValue}"/> property or field:
    /// each value is read from a column, keyed by the column's name.
    /// </summary>
    /// <typeparam name="TValue">The type of the values.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Scores</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, TValue> Map<TValue>(Expression<Func<T, IDictionary<string, TValue>>> expression)
        => ExpressionAutoMapper.GetOrCreateDictionaryMap<T, string, TValue>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="IReadOnlyDictionary{TKey, TValue}"/> property or field:
    /// each value is read from a column, keyed by the column's name.
    /// </summary>
    /// <typeparam name="TValue">The type of the values.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Scores</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, TValue> Map<TValue>(Expression<Func<T, IReadOnlyDictionary<string, TValue>>> expression)
        => ExpressionAutoMapper.GetOrCreateDictionaryMap<T, string, TValue>(this, expression.Body);

    /// <summary>
    /// Maps a <see cref="Dictionary{TKey, TValue}"/> property or field:
    /// each value is read from a column, keyed by the column's name.
    /// </summary>
    /// <typeparam name="TValue">The type of the values.</typeparam>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Scores</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, TValue> Map<TValue>(Expression<Func<T, Dictionary<string, TValue>>> expression)
        => ExpressionAutoMapper.GetOrCreateDictionaryMap<T, string, TValue>(this, expression.Body);

    /// <summary>
    /// Maps a <see cref="StringDictionary"/> property or field:
    /// each value is read from a column, keyed by the column's name.
    /// </summary>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Scores</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, string> Map(Expression<Func<T, StringDictionary>> expression)
        => ExpressionAutoMapper.GetOrCreateDictionaryMap<T, string, string>(this, expression.Body);

    /// <summary>
    /// Maps a <see cref="NameValueCollection"/> property or field:
    /// each value is read from a column, keyed by the column's name.
    /// </summary>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Scores</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, string> Map(Expression<Func<T, NameValueCollection>> expression)
        => ExpressionAutoMapper.GetOrCreateDictionaryMap<T, string, string>(this, expression.Body);

    /// <summary>
    /// Maps an <see cref="ExpandoObject"/> property or field:
    /// each value is read from a column, keyed by the column's name.
    /// </summary>
    /// <param name="expression">An expression reading the property or field, such as <c>p => p.Scores</c>.</param>
    /// <returns>The map of the property or field, to choose its columns and configure its values.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public DictionaryMap<string, object> Map(Expression<Func<T, ExpandoObject>> expression)
        => ExpressionAutoMapper.GetOrCreateDictionaryMap<T, string, object>(this, expression.Body);

    /// <summary>
    /// Maps every member that auto-mapping would, from its type and attributes, except members this
    /// class map already maps. Members mapped afterwards with <c>Map(...)</c> replace the
    /// auto-mapped ones, so a class map can start from auto-mapping and change only what's different:
    /// <code>
    /// importer.Configuration.RegisterClassMap&lt;Employee&gt;(c =&gt;
    /// {
    ///     c.AutoMap();
    ///     c.Map(e =&gt; e.Name).WithColumnName("Full Name");
    /// });
    /// </code>
    /// </summary>
    /// <returns>The class map.</returns>
    /// <exception cref="ExcelMappingException"><typeparamref name="T"/>, or a member's type, can't be auto-mapped.</exception>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public ExcelClassMap<T> AutoMap()
    {
        AutoMapper.AutoMap(this);
        return this;
    }

    /// <summary>
    /// Configures the existing class map used to map multiple cells in a row to the properties and fields
    /// of an object.
    /// </summary>
    /// <param name="configure">Configures this class map.</param>
    /// <returns>The map that invoked this method.</returns>
    public ExcelClassMap<T> WithClassMap(Action<ExcelClassMap<T>> configure)
    {
        ThrowHelpers.ThrowIfNull(configure, nameof(configure));

        configure(this);
        return this;
    }

    /// <summary>
    /// Sets the new class map used to map multiple cells in a row to the properties and fields
    /// of an object.
    /// </summary>
    /// <param name="classMap">The new class map used.</param>
    /// <returns>The map that invoked this method.</returns>
    public ExcelClassMap<T> WithClassMap(ExcelClassMap<T> classMap)
    {
        ThrowHelpers.ThrowIfNull(classMap, nameof(classMap));
        if (classMap == this)
        {
            return this;
        }

        Members.Clear();
        foreach (var member in classMap.Members)
        {
            Members.Add(member);
        }

        return this;
    }
}
