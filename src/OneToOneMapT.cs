using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Readers;
using ExcelMapper.Fallbacks;

namespace ExcelMapper;

/// <summary>
/// Maps a member to an Excel column of type T with a one-to-one relationship.
/// </summary>
/// <typeparam name="T">The type of the member.</typeparam>
public class OneToOneMap<T> : IOneToOneMap, IValuePipeline<T>, IBindableMap
{
    /// <summary>
    /// Constructs a one-to-one map with the specified cell reader factory.
    /// </summary>
    /// <param name="readerFactory">The factory used to create cell readers for the mapped column.</param>
    public OneToOneMap(ICellReaderFactory readerFactory)
    {
        ThrowHelpers.ThrowIfNull(readerFactory, nameof(readerFactory));
        _readerFactory = readerFactory;
    }

    private ICellReaderFactory _readerFactory;

    /// <inheritdoc />
    public ICellReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set
        {
            ThrowHelpers.ThrowIfNull(value, nameof(value));
            _readerFactory = value;
            ReaderVersion++;
        }
    }

    /// <inheritdoc />
    public bool Optional { get; set; }

    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    /// <inheritdoc />
    public IValuePipeline Pipeline { get; } = new ValuePipeline<T>();

    /// <inheritdoc />
    public IList<Abstractions.ICellTransformer> Transformers => Pipeline.Transformers;

    /// <inheritdoc />
    public IList<ICellMapper> Mappers => Pipeline.Mappers;

    /// <inheritdoc />
    public IFallbackItem? EmptyFallback
    {
        get => Pipeline.EmptyFallback;
        set => Pipeline.EmptyFallback = value;
    }

    /// <inheritdoc />
    public IFallbackItem? InvalidFallback
    {
        get => Pipeline.InvalidFallback;
        set => Pipeline.InvalidFallback = value;
    }

    /// <inheritdoc />
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
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
    /// (see <see cref="BoundOneToOneMember{T}"/>) are resolved again.
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
    /// <see cref="BoundOneToOneMember{T}"/>), with the same result as reading it through a
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
        => ((ValuePipeline<T>)Pipeline).Map(sheet, rowIndex, readResult, PreserveFormatting, member);

    /// <inheritdoc/>
    BoundMember IBindableMap.Bind(ExcelPropertyMap property, ExcelSheet sheet) => new BoundOneToOneMember<T>(this, property, sheet);

    /// <inheritdoc/>
    MethodInfo? IBindableMap.GetArgumentMethod(Type parameterType)
        => parameterType == typeof(T) ? s_getArgumentMethod ??= new Func<BoundOneToOneMember<T>, T, T>(BoundOneToOneMember<T>.GetValueOrDefault).Method : null;

    private static MethodInfo? s_getArgumentMethod;

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the map.</typeparam>
    /// <param name="value">The value that will be assigned to the property or field if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public OneToOneMap<T> WithValueFallback(T? value)
        => this.WithFallbackItem(new FixedValueFallback(value));

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="T">The type of the map.</typeparam>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell is empty.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public OneToOneMap<T> WithEmptyFallback(T? fallbackValue)
        => this.WithEmptyFallbackItem(new FixedValueFallback(fallbackValue));

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the map.</typeparam>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public OneToOneMap<T> WithInvalidFallback(T? fallbackValue)
        => this.WithInvalidFallbackItem(new FixedValueFallback(fallbackValue));
}
