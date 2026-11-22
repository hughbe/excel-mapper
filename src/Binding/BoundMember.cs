using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Readers;

namespace ExcelMapper.Binding;

/// <summary>
/// A member of a class map bound to a sheet: it reads the member's value from the current row and
/// sets it. Bound members are created once per sheet and class map (see
/// <see cref="SheetView.GetBoundMembers(ExcelClassMap)"/>), so they can keep what was resolved for
/// the sheet, such as the column to read. They're used by one sheet, so they aren't shared between threads.
/// </summary>
internal abstract class BoundMember(ClassMapMember member)
{
    /// <summary>
    /// Gets the class map's member: the property or field, and its map.
    /// </summary>
    protected ClassMapMember Member { get; } = member;

    /// <summary>
    /// Reads the member's value from the current row and sets it on <paramref name="instance"/>.
    /// </summary>
    /// <param name="instance">The object whose member is set.</param>
    /// <param name="sheet">The sheet being read.</param>
    /// <param name="rowIndex">The index of the row being read.</param>
    /// <param name="reader">The reader, on the row being read.</param>
    /// <param name="errors">The errors being collected, if any. A member whose value caused an
    /// error isn't set, so that its initial value is kept.</param>
    /// <returns>Whether the member was set.</returns>
    public virtual bool TryMap(object instance, SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        if (!TryGetValue(sheet, rowIndex, reader, errors, out var value))
        {
            return false;
        }

        Member.SetValue(instance, value!);
        return true;
    }

    /// <summary>
    /// Reads the member's value from the current row, for a constructor argument.
    /// </summary>
    /// <param name="sheet">The sheet being read.</param>
    /// <param name="rowIndex">The index of the row being read.</param>
    /// <param name="reader">The reader, on the row being read.</param>
    /// <param name="errors">The errors being collected, if any. Returns false if reading the value
    /// recorded an error.</param>
    /// <param name="value">The value read.</param>
    /// <returns>Whether there is a value.</returns>
    public abstract bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value);

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
    /// <param name="sheet">The sheet being read.</param>
    /// <param name="rowIndex">The index of the row being read.</param>
    /// <param name="reader">The reader, on the row being read.</param>
    /// <param name="errors">The errors being collected, if any. There is no value if reading it
    /// recorded an error.</param>
    public void ReadValue(SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        // Cleared first, in case reading throws.
        HasValue = false;
        HasValue = TryReadValue(sheet, rowIndex, reader, errors);
    }

    private protected virtual bool TryReadValue(SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
        => TryGetValue(sheet, rowIndex, reader, errors, out _value);

    /// <summary>
    /// The value read by <see cref="ReadValue"/>, as an object.
    /// </summary>
    public virtual object? Value => _value;

    /// <summary>
    /// Sets the value read by <see cref="ReadValue"/> on <paramref name="instance"/>.
    /// </summary>
    public virtual void SetValue(object instance) => Member.SetValue(instance, _value!);

    /// <summary>
    /// Returns the value read by <see cref="ReadValue"/>, or <paramref name="defaultValue"/>. Called
    /// by compiled constructors.
    /// </summary>
    public static object? GetValueOrDefault(BoundMember member, object? defaultValue)
        => member.HasValue ? member.Value : defaultValue;

    /// <summary>
    /// Whether reading a value added no errors to <paramref name="errors"/>: a value that caused
    /// an error isn't used.
    /// </summary>
    private protected static bool AddedNoErrors(List<ExcelMappingException>? errors, int errorCount)
        => errors is null || errors.Count == errorCount;
}

/// <summary>
/// Any map bound to a member, mapped through <see cref="IMap.TryGetValue"/>: for example nested
/// class maps, indexer maps, row numbers and custom maps.
/// </summary>
internal sealed class BoundMapMember(ClassMapMember member) : BoundMember(member)
{
    public override bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value)
    {
        var errorCount = errors?.Count ?? 0;
        if (Member.Map.TryGetValue(sheet, rowIndex, reader, Member.Member, errors, out value) && AddedNoErrors(errors, errorCount))
        {
            return true;
        }

        value = null;
        return false;
    }
}

/// <summary>
/// An <see cref="ICellsMap"/> bound to a member.
/// </summary>
internal sealed class BoundCellsMember(ICellsMap map, ClassMapMember member) : BoundMember(member)
{
    private ICellsReader? _cellsReader;
    private bool _resolved;

    public override bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value)
    {
        // Found on first use, inside the caller's error handling. Maps are frozen once rows are
        // read, so the reader stays right.
        if (!_resolved)
        {
            _cellsReader = map.GetCellsReader(sheet);
            _resolved = true;
        }

        var errorCount = errors?.Count ?? 0;
        if (map.TryGetValue(_cellsReader, sheet, rowIndex, reader, Member.Member, errors, out value) && AddedNoErrors(errors, errorCount))
        {
            return true;
        }

        value = null;
        return false;
    }
}

/// <summary>
/// A <see cref="ValueMap{T}"/> bound to a member: the value stays a <typeparamref name="T"/>
/// from the map to the member's setter.
/// </summary>
internal sealed class BoundValueMember<T>(ValueMap<T> map, ClassMapMember member) : BoundMember(member)
{
    // Created on first use: members passed to a constructor (such as get-only properties) are
    // never set.
    private Action<object, T>? _setter;

    // The member's column in the sheet (-1 if missing), found on first use (inside the caller's
    // error handling, as for unbound maps).
    private int _columnIndex = -1;
    private bool _resolved;

    // The value read by TryReadValue, kept typed.
    private T _value = default!;

    public override bool TryMap(object instance, SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        if (!TryRead(sheet, rowIndex, reader, errors, out var value))
        {
            return false;
        }

        (_setter ??= CompiledAccessors.GetSetter<T>(Member.Member))(instance, value);
        return true;
    }

    private bool TryRead(SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out T value)
    {
        if (!_resolved)
        {
            _columnIndex = map.FindColumn(sheet);
            _resolved = true;
        }

        var errorCount = errors?.Count ?? 0;
        return map.TryGetValue(_columnIndex, sheet, rowIndex, reader, Member.Member, errors, out value) && AddedNoErrors(errors, errorCount);
    }

    public override bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, out object? value)
    {
        var found = TryRead(sheet, rowIndex, reader, errors, out var typedValue);
        value = found ? typedValue : null;
        return found;
    }

    private protected override bool TryReadValue(SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
        => TryRead(sheet, rowIndex, reader, errors, out _value);

    public override object? Value => _value;

    public override void SetValue(object instance)
        => (_setter ??= CompiledAccessors.GetSetter<T>(Member.Member))(instance, _value);

    /// <summary>
    /// Returns the value read by <see cref="BoundMember.ReadValue"/>, or <paramref name="defaultValue"/>,
    /// without boxing. Called by compiled constructors.
    /// </summary>
    public static T GetValueOrDefault(BoundValueMember<T> member, T defaultValue)
        => member.HasValue ? member._value : defaultValue;
}
