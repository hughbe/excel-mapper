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
}

/// <summary>
/// A <see cref="OneToOneMap{T}"/> bound to a member: the value stays a <typeparamref name="T"/>
/// from the map to the member's setter.
/// </summary>
internal sealed class BoundOneToOneMember<T> : BoundMember
{
    private readonly OneToOneMap<T> _map;
    private readonly MemberInfo _member;
    private readonly Action<object, T> _setter;

    // The cell reader for the sheet, resolved on first use (inside the caller's error handling, as
    // for unbound maps) and again if the map's reader factory changes.
    private ICellReader? _cellReader;
    private int _readerVersion = -1;

    public BoundOneToOneMember(OneToOneMap<T> map, ExcelPropertyMap property, ExcelSheet sheet)
    {
        _map = map;
        _member = property.Member;
        _setter = CompiledAccessors.GetSetter<T>(property.Member);
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

        _setter(instance, value);
        return true;
    }
}
