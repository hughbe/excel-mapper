using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Utilities;

/// <summary>
/// A map that can be bound to a member of a class for a sheet, to map rows without per-cell
/// lookups or boxing. See <see cref="BoundMember"/>.
/// </summary>
internal interface IBindableMap
{
    BoundMember Bind(ExcelPropertyMap property, ExcelSheet sheet);
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

    public BoundOneToOneMember(OneToOneMap<T> map, ExcelPropertyMap property, ExcelSheet sheet)
    {
        _map = map;
        _member = property.Member;
    }

    public override bool TryMap(object instance, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        if (_readerVersion != _map.ReaderVersion)
        {
            _cellReader = _map.GetCellReader(sheet);
            _readerVersion = _map.ReaderVersion;
        }

        var errorCount = errors?.Count ?? 0;
        if (!_map.TryGetValue(_cellReader, sheet, rowIndex, reader, _member, out var value) ||
            (errors is not null && errors.Count != errorCount))
        {
            return false;
        }

        (_setter ??= CompiledAccessors.GetSetter<T>(_member))(instance, value);
        return true;
    }

    public override bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value)
    {
        if (_readerVersion != _map.ReaderVersion)
        {
            _cellReader = _map.GetCellReader(sheet);
            _readerVersion = _map.ReaderVersion;
        }

        var errorCount = errors?.Count ?? 0;
        if (_map.TryGetObjectValue(_cellReader, sheet, rowIndex, reader, _member, out value) &&
            (errors is null || errors.Count == errorCount))
        {
            return true;
        }

        value = null;
        return false;
    }
}
