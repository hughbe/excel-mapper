using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Readers;

namespace ExcelMapper;

/// <summary>
/// Maps a member to an Excel column of type T with a one-to-one relationship.
/// </summary>
/// <typeparam name="T">The type of the member.</typeparam>
public class ValueMap<T> : ValuePipeline<T, ValueMap<T>>, IValueMap, IBindableMap
{
    /// <summary>
    /// Constructs a one-to-one map with the specified cell reader factory.
    /// </summary>
    /// <param name="readerFactory">The factory used to create cell readers for the mapped column.</param>
    internal ValueMap(ICellReaderFactory readerFactory)
    {
        ThrowHelpers.ThrowIfNull(readerFactory, nameof(readerFactory));
        _readerFactory = readerFactory;
    }

    private ICellReaderFactory _readerFactory;

    /// <inheritdoc />
    internal ICellReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set
        {
            ThrowHelpers.ThrowIfNull(value, nameof(value));
            _readerFactory = value;
            ReaderVersion++;
        }
    }

    ICellReaderFactory IValueMap.ReaderFactory
    {
        get => ReaderFactory;
        set => ReaderFactory = value;
    }

    /// <inheritdoc />
    public bool Optional { get; set; }

    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    /// <summary>
    /// Gets the map's value pipeline: the map itself, which converts the cell's value.
    /// </summary>
    ValuePipeline IToOneMap.Pipeline => this;

    /// <inheritdoc />
    bool IMap.TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        if (TryGetValue(GetCellReader(sheet), sheet, rowIndex, reader, member, out var value))
        {
            result = value!;
            return true;
        }

        result = null;
        return false;
    }

    /// <summary>
    /// Incremented when <see cref="ReaderFactory"/> changes, so that readers resolved for a sheet
    /// (see <see cref="BoundValueMember{T}"/>) are resolved again.
    /// </summary>
    internal int ReaderVersion { get; private set; }

    internal ICellReader? GetCellReader(ExcelSheet sheet)
        => sheet.GetCellReader(_readerFactory);

    /// <summary>
    /// Reads and maps the cell, returning the value as <typeparamref name="T"/> so that value types
    /// aren't boxed again on the way to the member.
    /// </summary>
    internal bool TryGetValue(ICellReader? cellReader, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, out T value)
    {
        if (cellReader == null || !cellReader.TryGetValue(reader, PreserveFormatting, out var readResult))
        {
            return NoColumn(sheet, rowIndex, member, out value);
        }

        value = Map(readResult, sheet, rowIndex, member);
        return true;
    }

    /// <summary>
    /// Reads and maps the cell in the given column, for a member whose column is known (see
    /// <see cref="BoundValueMember{T}"/>), with the same result as reading it through a
    /// <see cref="ColumnIndexReader"/>.
    /// </summary>
    internal bool TryGetValue(int columnIndex, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, out T value)
    {
        if (columnIndex >= reader.FieldCount)
        {
            return NoColumn(sheet, rowIndex, member, out value);
        }

        value = Map(new ReadCellResult(columnIndex, reader, PreserveFormatting), sheet, rowIndex, member);
        return true;
    }

    private bool NoColumn(ExcelSheet sheet, int rowIndex, MemberInfo? member, out T value)
    {
        if (Optional)
        {
            value = default!;
            return false;
        }

        throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _readerFactory, member);
    }

    private T Map(ReadCellResult readResult, ExcelSheet sheet, int rowIndex, MemberInfo? member)
        => Map(sheet, rowIndex, readResult, PreserveFormatting, member);

    /// <inheritdoc/>
    BoundMember IBindableMap.Bind(ExcelPropertyMap property, ExcelSheet sheet) => new BoundValueMember<T>(this, property, sheet);

    /// <inheritdoc/>
    MethodInfo? IBindableMap.GetArgumentMethod(Type parameterType)
        => parameterType == typeof(T) ? s_getArgumentMethod ??= new Func<BoundValueMember<T>, T, T>(BoundValueMember<T>.GetValueOrDefault).Method : null;

    private static MethodInfo? s_getArgumentMethod;

    /// <summary>
    /// Reads the member from the column with the given name, ignoring case.
    /// </summary>
    /// <param name="columnName">The name of the column.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnName(string columnName)
        => WithColumnName(columnName, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Reads the member from the column with the given name.
    /// </summary>
    /// <param name="columnName">The name of the column.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnName(string columnName, StringComparison comparison)
        => WithReaderFactory(new ColumnNameReaderFactory(columnName, comparison));

    /// <summary>
    /// Reads the member from the first of the given columns that the sheet has.
    /// </summary>
    /// <param name="columnNames">The names of the columns, in order of preference.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNameMatching(params string[] columnNames)
        => WithReaderFactory(new ColumnNamesReaderFactory(columnNames));

    /// <summary>
    /// Reads the member from the first column whose name matches <paramref name="predicate"/>.
    /// </summary>
    /// <param name="predicate">Whether a column name matches.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNameMatching(Func<string, bool> predicate)
        => WithReaderFactory(new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(predicate)));

    /// <summary>
    /// Reads the member from the first column whose name matches <paramref name="regex"/>.
    /// </summary>
    /// <param name="regex">The regular expression column names are matched against.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNameMatching(Regex regex)
        => WithReaderFactory(new ColumnsMatchingReaderFactory(new RegexColumnMatcher(regex)));

    /// <summary>
    /// Reads the member from the column at the given zero-based index.
    /// </summary>
    /// <param name="columnIndex">The zero-based index of the column.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnIndex(int columnIndex)
        => WithReaderFactory(new ColumnIndexReaderFactory(columnIndex));

    /// <summary>
    /// Reads the member from the column with the given letter, such as "A" or "AB".
    /// </summary>
    /// <param name="columnLetter">The letter of the column.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnLetter(string columnLetter)
        => WithColumnIndex(ExcelRange.ParseColumnLetters(columnLetter, nameof(columnLetter)));

    /// <summary>
    /// Reads the member from the first column <paramref name="selector"/> chooses for each sheet.
    /// </summary>
    /// <param name="selector">The column selector.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumns(IColumnSelector selector)
    {
        ThrowHelpers.ThrowIfNull(selector, nameof(selector));
        return WithReaderFactory(new ColumnSelectorReaderFactory(selector));
    }

    /// <summary>
    /// Makes the member optional: if its column is missing, the member keeps its value instead of
    /// an error being thrown.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> MakeOptional()
    {
        Optional = true;
        return this;
    }

    /// <summary>
    /// Reads the cell's value as the text Excel shows, with its number format applied.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> MakePreserveFormatting()
    {
        PreserveFormatting = true;
        return this;
    }

    private ValueMap<T> WithReaderFactory(ICellReaderFactory readerFactory)
    {
        ReaderFactory = readerFactory;
        return this;
    }
}
