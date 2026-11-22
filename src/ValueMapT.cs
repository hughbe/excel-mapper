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
public class ValueMap<T> : ValuePipeline<T, ValueMap<T>>, IValueMap
{
    internal ValueMap(Columns columns)
    {
        _columns = columns;
    }

    private Columns _columns;

    /// <summary>
    /// Gets or sets the columns the member is read from: the first one the sheet has.
    /// </summary>
    internal Columns Columns
    {
        get => _columns;
        set
        {
            _columns = value;
            ReaderVersion++;
        }
    }

    Columns IColumnsMap.Columns
    {
        get => Columns;
        set => Columns = value;
    }

    /// <inheritdoc />
    public bool Optional { get; set; }

    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    /// <summary>
    /// Gets the map's value pipeline: the map itself, which converts the cell's value.
    /// </summary>
    ValuePipeline IColumnsMap.Pipeline => this;

    /// <inheritdoc />
    bool IMap.TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? result)
    {
        if (TryGetValue(FindColumn(sheet), sheet, rowIndex, reader, member, errors, out var value))
        {
            result = value!;
            return true;
        }

        result = null;
        return false;
    }

    /// <summary>
    /// Incremented when <see cref="Columns"/> changes, so that the column found for a sheet (see
    /// <see cref="BoundValueMember{T}"/>) is found again.
    /// </summary>
    internal int ReaderVersion { get; private set; }

    /// <summary>
    /// Finds the member's column in the sheet, or -1 if the sheet doesn't have it.
    /// </summary>
    internal int FindColumn(ExcelSheet sheet) => sheet.FindColumn(_columns);

    /// <summary>
    /// Reads and maps the cell in the given column (-1 if the sheet doesn't have the member's
    /// column), returning the value as <typeparamref name="T"/> so that value types aren't boxed
    /// again on the way to the member.
    /// </summary>
    internal bool TryGetValue(int columnIndex, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, out T value)
    {
        if (columnIndex == -1 || columnIndex >= reader.FieldCount)
        {
            if (Optional)
            {
                value = default!;
                return false;
            }

            throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _columns, member);
        }

        value = Map(sheet, rowIndex, new ReadCellResult(columnIndex, reader, PreserveFormatting), member, errors);
        return true;
    }


    /// <inheritdoc/>
    BoundMember IValueMap.Bind(ExcelPropertyMap property, ExcelSheet sheet) => new BoundValueMember<T>(this, property, sheet);

    /// <inheritdoc/>
    MethodInfo? IValueMap.GetArgumentMethod(Type parameterType)
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
        => WithColumns(new ColumnNameColumns(columnName, comparison));

    /// <summary>
    /// Reads the member from the first of the given columns that the sheet has.
    /// </summary>
    /// <param name="columnNames">The names of the columns, in order of preference.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNameMatching(params string[] columnNames)
        => WithColumns(new ColumnNamesColumns(columnNames));

    /// <summary>
    /// Reads the member from the first column whose name matches <paramref name="predicate"/>.
    /// </summary>
    /// <param name="predicate">Whether a column name matches.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNameMatching(Func<string, bool> predicate)
        => WithColumns(new MatchingColumns(predicate));

    /// <summary>
    /// Reads the member from the first column whose name matches <paramref name="regex"/>.
    /// </summary>
    /// <param name="regex">The regular expression column names are matched against.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNameMatching(Regex regex)
        => WithColumns(new MatchingColumns(regex));

    /// <summary>
    /// Reads the member from the column at the given zero-based index.
    /// </summary>
    /// <param name="columnIndex">The zero-based index of the column.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnIndex(int columnIndex)
        => WithColumns(new ColumnIndexColumns(columnIndex));

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
        return WithColumns(new SelectorColumns(selector));
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

    private ValueMap<T> WithColumns(Columns columns)
    {
        Columns = columns;
        return this;
    }
}
