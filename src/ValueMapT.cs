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
public sealed class ValueMap<T> : ValuePipeline<T, ValueMap<T>>, IColumnsMap, IBindableMap
{
    internal ValueMap(Columns columns)
    {
        Settings = new ColumnSettings(columns);
    }

    /// <summary>
    /// Gets or sets the columns the map reads.
    /// </summary>
    internal Columns Columns
    {
        get => Settings.Columns;
        set => Settings.Columns = value;
    }

    /// <summary>
    /// Gets where the map reads its cells, and how.
    /// </summary>
    internal ColumnSettings Settings { get; }

    ColumnSettings IColumnsMap.Settings => Settings;

    /// <summary>
    /// Gets whether the member's columns can be missing from the sheet: the member then
    /// keeps its value, rather than mapping the row being an error.
    /// </summary>
    public bool Optional => Settings.Optional;

    /// <summary>
    /// Gets whether cells are read as the text Excel shows, with their number format applied.
    /// </summary>
    public bool PreserveFormatting => Settings.PreserveFormatting;

    CellPipeline IColumnsMap.Pipeline => Pipeline;

    /// <inheritdoc />
    void IMap.Freeze()
    {
        Settings.Freeze();
        Pipeline.Freeze();
    }

    /// <inheritdoc />
    bool IMap.TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? result)
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
    /// Finds the member's column in the sheet, or -1 if the sheet doesn't have it.
    /// </summary>
    internal int FindColumn(SheetView sheet) => sheet.FindColumn(Settings.Columns);

    /// <summary>
    /// Reads and maps the cell in the given column (-1 if the sheet doesn't have the member's
    /// column), returning the value as <typeparamref name="T"/> so that value types aren't boxed
    /// again on the way to the member.
    /// </summary>
    internal bool TryGetValue(int columnIndex, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, out T value)
    {
        if (columnIndex == -1 || columnIndex >= reader.FieldCount)
        {
            value = default!;
            return Settings.MissingColumn(sheet, rowIndex, member);
        }

        value = Pipeline.Map(sheet, rowIndex, new ReadCellResult(columnIndex, reader, PreserveFormatting), member, errors);
        return true;
    }


    /// <inheritdoc/>
    BoundMember<TOwner> IBindableMap.Bind<TOwner>(ClassMapMember member) => new BoundValueMember<TOwner, T>(this, member);

    /// <inheritdoc/>
    MethodInfo? IBindableMap.GetArgumentMethod<TOwner>(Type parameterType)
        => parameterType == typeof(T) ? new Func<BoundMember<TOwner, T>, T, T>(BoundMember<TOwner, T>.GetValueOrDefault).Method : null;

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
    /// Reads the member from the first of the given columns that the sheet has, as
    /// <see cref="ExcelColumnNamesAttribute"/> does.
    /// </summary>
    /// <param name="columnNames">The names of the columns, in order of preference.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNames(params string[] columnNames)
        => WithColumns(new ColumnNamesColumns(columnNames));

    /// <summary>
    /// Reads the member from the first of the given columns that the sheet has, as
    /// <see cref="ExcelColumnNamesAttribute"/> does.
    /// </summary>
    /// <param name="columnNames">The names of the columns, in order of preference.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNames(IEnumerable<string> columnNames)
        => WithColumns(new ColumnNamesColumns(columnNames));

    /// <summary>
    /// Reads the member from the first of the given columns that the sheet has, as
    /// <see cref="ExcelColumnNamesAttribute"/> does.
    /// </summary>
    /// <param name="columnNames">The names of the columns, in order of preference.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> WithColumnNames(IEnumerable<string> columnNames, StringComparison comparison)
        => WithColumns(new ColumnNamesColumns(columnNames, comparison));

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
        => WithColumns(new SelectorColumns(selector));

    /// <summary>
    /// Makes the member optional: if its column is missing, the member keeps its value instead of
    /// an error being thrown.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> MakeOptional()
    {
        Settings.Optional = true;
        return this;
    }

    /// <summary>
    /// Reads the cell's value as the text Excel shows, with its number format applied.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public ValueMap<T> MakePreserveFormatting()
    {
        Settings.PreserveFormatting = true;
        return this;
    }

    private ValueMap<T> WithColumns(Columns columns)
    {
        Columns = columns;
        return this;
    }
}
