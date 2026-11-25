using ExcelMapper.Readers;
using System.Text.RegularExpressions;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Builders;

namespace ExcelMapper;

/// <summary>
/// A map that reads one or more values from one or more cells and maps these values to the type of the
/// property or field. This is used to map IDictionary properties and fields.
/// </summary>
/// <typeparam name="TKey">The key type of the IDictionary property or field.</typeparam>
/// <typeparam name="TValue">The value type of the IDictionary property or field.</typeparam>
public sealed class DictionaryMap<TKey, TValue> : ICellsMap where TKey : notnull
{
    /// <summary>
    /// Constructs a map that reads values from one or more cells and maps them to the values of the
    /// dictionary assigned to the property or field.
    /// </summary>
    /// <param name="columns">The columns to read.</param>
    /// <param name="builder">Builds the dictionary.</param>
    /// <param name="keyMappers">Convert column names to keys that aren't strings: the built-in
    /// conversion of <typeparamref name="TKey"/>.</param>
    internal DictionaryMap(Columns columns, DictionaryBuilder<TKey, TValue> builder, ICellMapper<TKey>[] keyMappers)
    {
        Settings = new ColumnSettings(columns);
        Builder = builder;
        _keyMappers = keyMappers;
    }

    private readonly ICellMapper<TKey>[] _keyMappers;

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

    /// <summary>
    /// Gets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    internal ElementMap<TValue> Values { get; } = new ElementMap<TValue>();

    CellPipeline IColumnsMap.Pipeline => Values.Pipeline;

    /// <inheritdoc />
    void IMap.Freeze()
    {
        Settings.Freeze();
        Values.Pipeline.Freeze();
    }

    /// <summary>
    /// How the member's dictionary is built from its values.
    /// </summary>
    internal DictionaryBuilder<TKey, TValue> Builder { get; }

    /// <inheritdoc/>
    bool IMap.TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? value)
    {
        return TryGetValue(GetCellsReader(sheet), sheet, rowIndex, reader, member, errors, out value);
    }

    ICellsReader? ICellsMap.GetCellsReader(SheetView sheet) => GetCellsReader(sheet);

    private ICellsReader? GetCellsReader(SheetView sheet)
        => sheet.GetCellsReader(Settings.Columns);

    bool ICellsMap.TryGetValue(ICellsReader? cellsReader, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, out object? value)
        => TryGetValue(cellsReader, sheet, rowIndex, reader, member, errors, out value);

    private bool TryGetValue(ICellsReader? cellsReader, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? value)
    {
        var heading = sheet.GetHeadingForKeys();
        if (!Settings.TryStart(cellsReader, sheet, rowIndex, reader, member, out var count))
        {
            value = default;
            return false;
        }

        // String and object keys are the columns' names. Keys of other types are converted once for
        // the sheet.
        var keys = IsNameKey ? null : sheet.GetDictionaryKeys(this, s_createKeys);

        var items = Builder.Begin(count);
        try
        {
            while (cellsReader.TryGetNext(out var valueResult))
            {
                // The key is the column's name.
                var columnName = heading.GetColumnName(valueResult.ColumnIndex);
                var key = keys is null ? (TKey)(object)columnName : keys.Get(this, heading, sheet, valueResult.ColumnIndex, member);
                var elementValue = Values.Pipeline.Map(sheet, rowIndex, valueResult, member, errors);
                try
                {
                    Builder.Add(items, key, elementValue);
                }
                catch (ArgumentException exception)
                {
                    // Dictionaries reject a key they already have: two columns with the same name.
                    throw DuplicateKey(heading, sheet, valueResult.ColumnIndex, columnName, member, exception);
                }
            }
        }
        finally
        {
            cellsReader.Reset();
        }

        value = Builder.End(items);
        return true;
    }

    /// <summary>
    /// Whether the keys are the columns' names, rather than converted from them.
    /// </summary>
    internal static bool IsNameKey => typeof(TKey) == typeof(string) || typeof(TKey) == typeof(object);

    private static readonly Func<Keys> s_createKeys = static () => new Keys();

    /// <summary>
    /// The keys of a sheet's columns, each converted the first time a row reads it.
    /// </summary>
    private sealed class Keys
    {
        private ExcelHeading? _heading;
        private TKey[] _keys = [];
        private bool[] _converted = [];

        public TKey Get(DictionaryMap<TKey, TValue> map, ExcelHeading heading, SheetView sheet, int columnIndex, MemberInfo? member)
        {
            if (!ReferenceEquals(heading, _heading))
            {
                _heading = heading;
                _keys = new TKey[heading.ColumnNames.Count];
                _converted = new bool[heading.ColumnNames.Count];
            }

            var index = columnIndex - heading.ColumnOffset;
            if (!_converted[index])
            {
                _keys[index] = map.ConvertKey(heading, sheet, columnIndex, member);
                _converted[index] = true;
            }

            return _keys[index];
        }
    }

    // A column's heading cell as a key of another type, such as an int, enum or date, converted as
    // a cell of that type is: a number or date in the heading as itself, and text with the
    // importer's culture.
    private TKey ConvertKey(ExcelHeading heading, SheetView sheet, int columnIndex, MemberInfo? member)
    {
        var cell = new ReadCellResult(columnIndex, heading.GetColumnValue(columnIndex))
        {
            Culture = sheet.Culture,
        };
        var context = new CellContext(sheet, -1, columnIndex, member, null);

        var state = CellConversion.NotHandled;
        TKey result = default!;
        var mappers = _keyMappers;
        for (var i = 0; i < mappers.Length; i++)
        {
            object? error = null;
            switch (mappers[i].Convert(cell, in context, out var key, ref error))
            {
                case CellConversion.Success:
                    return key;
                case CellConversion.SuccessUnlessReplaced:
                    state = CellConversion.SuccessUnlessReplaced;
                    result = key;
                    break;
            }
        }

        if (state == CellConversion.SuccessUnlessReplaced)
        {
            return result;
        }

        var columnName = heading.GetColumnName(columnIndex);
        throw new ExcelMappingException($"The name \"{columnName}\" of column {ExcelRange.GetColumnLetters(columnIndex)} can't be converted to {TypeNames.Get(typeof(TKey))}, the type of the dictionary's keys", sheet, -1, -1, null)
        {
            Member = member,
        };
    }

    private static ExcelMappingException DuplicateKey(ExcelHeading heading, SheetView sheet, int columnIndex, string columnName, MemberInfo? member, ArgumentException exception)
    {
        // The earlier column with the same name, which the dictionary already has: exactly the
        // same, or for a dictionary that ignores case, the same ignoring case.
        var firstColumnIndex = FindEarlierColumn(heading, columnIndex, columnName, StringComparison.Ordinal);
        if (firstColumnIndex == -1)
        {
            firstColumnIndex = FindEarlierColumn(heading, columnIndex, columnName, StringComparison.OrdinalIgnoreCase);
        }

        // Otherwise two names gave the same key, such as "1" and "01" for int keys.
        var message = firstColumnIndex == -1
            ? $"The name \"{columnName}\" of column {ExcelRange.GetColumnLetters(columnIndex)} gives a key the dictionary already has"
            : $"Columns {ExcelRange.GetColumnLetters(firstColumnIndex)} and {ExcelRange.GetColumnLetters(columnIndex)} are both named \"{columnName}\", so they can't both be keys of the dictionary";
        return new ExcelMappingException(message, sheet, -1, -1, exception, " Choose its columns with WithColumnNames or [ExcelColumnNames], or rename one of the columns.")
        {
            Member = member,
        };
    }

    private static int FindEarlierColumn(ExcelHeading heading, int columnIndex, string columnName, StringComparison comparison)
    {
        for (var i = 0; i + heading.ColumnOffset < columnIndex && i < heading.ColumnNames.Count; i++)
        {
            if (string.Equals(heading.ColumnNames[i], columnName, comparison))
            {
                return i + heading.ColumnOffset;
            }
        }

        return -1;
    }

    /// <summary>
    /// Configures how each value is converted from its cell, with the value methods a single
    /// member has: for example <c>.WithValueMap(v =&gt; v.WithTrim().WithEmptyFallback(0))</c>.
    /// </summary>
    /// <param name="configure">Configures the values' map.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithValueMap(Action<ElementMap<TValue>> configure)
    {
        ThrowHelpers.ThrowIfNull(configure, nameof(configure));
        Settings.ThrowIfFrozen();

        configure(Values);
        return this;
    }

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnNames(params string[] columnNames)
        => WithColumns(new ColumnNamesColumns(columnNames));

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnNames(IEnumerable<string> columnNames)
        => WithColumns(new ColumnNamesColumns(columnNames));

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnNames(IEnumerable<string> columnNames, StringComparison comparison)
        => WithColumns(new ColumnNamesColumns(columnNames, comparison));

    /// <summary>
    /// Reads the dictionary from every column whose name matches <paramref name="predicate"/>.
    /// </summary>
    /// <param name="predicate">Whether a column name matches.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnsMatching(Func<string, bool> predicate)
        => WithColumns(new MatchingColumns(predicate));

    /// <summary>
    /// Reads the dictionary from every column whose name matches <paramref name="regex"/>.
    /// </summary>
    /// <param name="regex">The regular expression column names are matched against.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnsMatching(Regex regex)
        => WithColumns(new MatchingColumns(regex));

    /// <summary>
    /// Reads the dictionary from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnIndices(params int[] columnIndices)
        => WithColumns(new ColumnIndicesColumns(columnIndices));

    /// <summary>
    /// Reads the dictionary from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnIndices(IEnumerable<int> columnIndices)
        => WithColumns(new ColumnIndicesColumns(columnIndices));

    /// <summary>
    /// Reads the dictionary from the columns <paramref name="selector"/> chooses for each sheet, in order.
    /// </summary>
    /// <param name="selector">The column selector.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumns(IColumnSelector selector)
        => WithColumns(new SelectorColumns(selector));

    /// <summary>
    /// Makes the member optional: if its columns are missing, the member keeps its value instead of
    /// an error being thrown.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> MakeOptional()
    {
        Settings.Optional = true;
        return this;
    }

    /// <summary>
    /// Reads the cells' values as the text Excel shows, with their number formats applied.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> MakePreserveFormatting()
    {
        Settings.PreserveFormatting = true;
        return this;
    }

    private DictionaryMap<TKey, TValue> WithColumns(Columns columns)
    {
        Columns = columns;
        return this;
    }
}
