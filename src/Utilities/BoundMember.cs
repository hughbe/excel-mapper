using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// A map that can be bound to a member of a class for a sheet, to map rows without per-cell
/// lookups or boxing. See <see cref="BoundMember"/>.
/// </summary>
internal interface IBindableMap
{
    BoundMember Bind(ExcelPropertyMap property, ExcelSheet sheet);

    /// <summary>
    /// Gets a static method <c>(TBoundMember member, TParameter defaultValue) → TParameter</c> that
    /// returns the value a bound member of this map read for a constructor parameter, without
    /// boxing. Null if the parameter's type isn't the map's type.
    /// </summary>
    MethodInfo? GetArgumentMethod(Type parameterType);
}

/// <summary>
/// A member of a class map bound to a sheet: it reads the member's value from the current row and
/// sets it. Bound members are created once per sheet and class map (see
/// <see cref="ExcelSheet.GetBoundMembers(ExcelClassMap)"/>), so they can keep what was resolved for
/// the sheet, such as the column to read. They're used by one sheet, so they aren't shared between threads.
/// </summary>
internal abstract class BoundMember
{
    /// <summary>
    /// Reads the member's value from the current row and sets it on <paramref name="instance"/>.
    /// </summary>
    /// <param name="errors">The errors being collected, if any. A member whose value caused an
    /// error isn't set, so that its initial value is kept.</param>
    /// <returns>Whether the member was set.</returns>
    public abstract bool TryMap(object instance, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors);

    /// <summary>
    /// Reads the member's value from the current row, for a constructor argument.
    /// </summary>
    /// <param name="errors">The errors being collected, if any. Returns false if reading the value
    /// recorded an error.</param>
    /// <returns>Whether there is a value.</returns>
    public abstract bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value);

    // The value read by ReadValue, for types created through a constructor: every member is read
    // before the object exists, and then passed to the constructor or set.
    private object? _value;

    /// <summary>
    /// Whether <see cref="ReadValue"/> read a value from the current row.
    /// </summary>
    public bool HasValue { get; private set; }

    /// <summary>
    /// Reads the member's value from the current row and keeps it, for a type created through a
    /// constructor. The value is passed to the constructor (see
    /// <see cref="CompiledAccessors.GetConstructor(ConstructorInfo, BoundConstructorArgument[])"/>),
    /// or set with <see cref="SetValue"/>.
    /// </summary>
    /// <param name="errors">The errors being collected, if any. There is no value if reading it
    /// recorded an error.</param>
    public void ReadValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        // Cleared first, in case reading throws.
        HasValue = false;
        HasValue = TryReadValue(sheet, rowIndex, reader, errors);
    }

    private protected virtual bool TryReadValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
        => TryGetValue(sheet, rowIndex, reader, errors, out _value);

    /// <summary>
    /// The value read by <see cref="ReadValue"/>, as an object.
    /// </summary>
    public virtual object? Value => _value;

    /// <summary>
    /// Sets the value read by <see cref="ReadValue"/> on <paramref name="instance"/>.
    /// </summary>
    public virtual void SetValue(object instance, ExcelPropertyMap property) => property.SetValueFactory(instance, _value!);

    /// <summary>
    /// Returns the value read by <see cref="ReadValue"/>, or <paramref name="defaultValue"/>. Called
    /// by compiled constructors.
    /// </summary>
    public static object? GetValueOrDefault(BoundMember member, object? defaultValue)
        => member.HasValue ? member.Value : defaultValue;
}

/// <summary>
/// Any map bound to a member, mapped through <see cref="IMap.TryGetValue"/>: for example nested
/// class maps, indexer maps, row numbers and custom maps.
/// </summary>
internal sealed class BoundMapMember(ExcelPropertyMap property) : BoundMember
{
    public override bool TryMap(object instance, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        if (!TryGetValue(sheet, rowIndex, reader, errors, out var value))
        {
            return false;
        }

        property.SetValueFactory(instance, value!);
        return true;
    }

    public override bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value)
    {
        var errorCount = errors?.Count ?? 0;
        if (property.Map.TryGetValue(sheet, rowIndex, reader, property.Member, out value) &&
            (errors is null || errors.Count == errorCount))
        {
            return true;
        }

        value = null;
        return false;
    }
}

/// <summary>
/// A map that reads several cells (a collection or dictionary) and can be bound to a member: its
/// cells reader is resolved once for the sheet.
/// </summary>
internal interface ICellsMap : IMap
{
    int ReaderVersion { get; }

    ICellsReader? GetCellsReader(ExcelSheet sheet);

    /// <summary>
    /// Creates the map's collection factory for one binding, or null if the factory can't be copied
    /// (a custom factory), in which case the map's own factory is used under a lock.
    /// </summary>
    object? CreateBindingFactory();

    bool TryGetValue(ICellsReader? cellsReader, object? bindingFactory, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, out object? value);
}

/// <summary>
/// An <see cref="ICellsMap"/> bound to a member.
/// </summary>
internal sealed class BoundCellsMember(ICellsMap map, ExcelPropertyMap property) : BoundMember
{
    private readonly object? _factory = map.CreateBindingFactory();
    private ICellsReader? _cellsReader;
    private int _readerVersion = -1;

    public override bool TryMap(object instance, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        if (!TryGetValue(sheet, rowIndex, reader, errors, out var value))
        {
            return false;
        }

        property.SetValueFactory(instance, value!);
        return true;
    }

    public override bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value)
    {
        if (_readerVersion != map.ReaderVersion)
        {
            _cellsReader = map.GetCellsReader(sheet);
            _readerVersion = map.ReaderVersion;
        }

        var errorCount = errors?.Count ?? 0;
        if (map.TryGetValue(_cellsReader, _factory, sheet, rowIndex, reader, property.Member, out value) &&
            (errors is null || errors.Count == errorCount))
        {
            return true;
        }

        value = null;
        return false;
    }
}

/// <summary>
/// A <see cref="OneToOneMap{T}"/> bound to a member: the value stays a <typeparamref name="T"/>
/// from the map to the member's setter.
/// </summary>
internal sealed class BoundOneToOneMember<T> : BoundMember
{
    private readonly OneToOneMap<T> _map;
    private readonly MemberInfo _member;
    // Created on first use: members passed to a constructor (such as get-only properties) are
    // never set.
    private Action<object, T>? _setter;

    // The cell reader for the sheet, resolved on first use (inside the caller's error handling, as
    // for unbound maps) and again if the map's reader factory changes.
    private ICellReader? _cellReader;
    private int _readerVersion = -1;

    // The column, if the reader is a plain column reader: the cell is then read directly.
    private int _columnIndex = -1;

    public BoundOneToOneMember(OneToOneMap<T> map, ExcelPropertyMap property, ExcelSheet sheet)
    {
        _map = map;
        _member = property.Member;
    }

    // The value read by TryReadValue, kept typed.
    private T _value = default!;

    public override bool TryMap(object instance, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        if (!TryRead(sheet, rowIndex, reader, errors, out var value))
        {
            return false;
        }

        (_setter ??= CompiledAccessors.GetSetter<T>(_member))(instance, value);
        return true;
    }

    private bool TryRead(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out T value)
    {
        if (_readerVersion != _map.ReaderVersion)
        {
            Resolve(sheet);
        }

        var errorCount = errors?.Count ?? 0;
        var found = _columnIndex >= 0
            ? _map.TryGetValue(_columnIndex, sheet, rowIndex, reader, _member, out value)
            : _map.TryGetValue(_cellReader, sheet, rowIndex, reader, _member, out value);
        return found && (errors is null || errors.Count == errorCount);
    }

    private void Resolve(ExcelSheet sheet)
    {
        _cellReader = _map.GetCellReader(sheet);
        _columnIndex = _cellReader is ColumnIndexReader columnIndexReader ? columnIndexReader.ColumnIndex : -1;
        _readerVersion = _map.ReaderVersion;
    }

    public override bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value)
    {
        var found = TryRead(sheet, rowIndex, reader, errors, out var typedValue);
        value = found ? typedValue : null;
        return found;
    }

    private protected override bool TryReadValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
        => TryRead(sheet, rowIndex, reader, errors, out _value);

    public override object? Value => _value;

    public override void SetValue(object instance, ExcelPropertyMap property)
        => (_setter ??= CompiledAccessors.GetSetter<T>(_member))(instance, _value);

    /// <summary>
    /// Returns the value read by <see cref="BoundMember.ReadValue"/>, or <paramref name="defaultValue"/>,
    /// without boxing. Called by compiled constructors.
    /// </summary>
    public static T GetValueOrDefault(BoundOneToOneMember<T> member, T defaultValue)
        => member.HasValue ? member._value : defaultValue;
}
