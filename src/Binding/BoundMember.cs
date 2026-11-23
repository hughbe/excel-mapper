using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Readers;

namespace ExcelMapper.Binding;

/// <summary>
/// A member of a class map bound to a sheet: it reads the member's value from the current row and
/// sets it on a <typeparamref name="TOwner"/>, the class map's type. Bound members are created once
/// per sheet and class map (see <see cref="SheetView.GetBoundMembers{TOwner}(ExcelClassMap{TOwner})"/>),
/// so they can keep what was resolved for the sheet, such as the column to read. They're used by one
/// sheet, so they aren't shared between threads.
/// </summary>
/// <remarks>
/// The instance is passed by reference, so a struct is set in place: rows of a value type are
/// never boxed.
/// </remarks>
internal abstract class BoundMember<TOwner>(ClassMapMember member)
{
    /// <summary>
    /// Gets the class map's member: the property or field, and its map.
    /// </summary>
    protected ClassMapMember Member { get; } = member;

    /// <summary>
    /// Reads the member's value from the current row and sets it on <paramref name="instance"/>.
    /// </summary>
    /// <param name="instance">The instance whose member is set.</param>
    /// <param name="sheet">The sheet being read.</param>
    /// <param name="rowIndex">The index of the row being read.</param>
    /// <param name="reader">The reader, on the row being read.</param>
    /// <param name="errors">The errors being collected, if any. A member whose value caused an
    /// error isn't set, so that its initial value is kept.</param>
    /// <returns>Whether the member was set.</returns>
    public abstract bool TryMap(ref TOwner instance, SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors);

    /// <summary>
    /// Whether <see cref="ReadValue"/> read a value from the current row.
    /// </summary>
    public bool HasValue { get; private set; }

    /// <summary>
    /// Reads the member's value from the current row and keeps it, for a type created through a
    /// constructor: every member is read before the instance exists. The value is passed to the
    /// constructor (see <see cref="CompiledAccessors.GetConstructor{TOwner}(ConstructorInfo, BoundConstructorArgument[])"/>),
    /// or set with <see cref="SetValue"/>.
    /// </summary>
    /// <param name="sheet">The sheet being read.</param>
    /// <param name="rowIndex">The index of the row being read.</param>
    /// <param name="reader">The reader, on the row being read.</param>
    /// <param name="errors">The errors being collected, if any. There is no value if reading it
    /// recorded an error.</param>
    public void ReadValue(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors)
    {
        // Cleared first, in case reading throws.
        HasValue = false;
        HasValue = TryReadValue(sheet, rowIndex, reader, errors);
    }

    // Reads the value for ReadValue and keeps it.
    private protected abstract bool TryReadValue(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors);

    /// <summary>
    /// The value read by <see cref="ReadValue"/>, as an object.
    /// </summary>
    public abstract object? Value { get; }

    /// <summary>
    /// Sets the value read by <see cref="ReadValue"/> on <paramref name="instance"/>.
    /// </summary>
    public abstract void SetValue(ref TOwner instance);

    /// <summary>
    /// Returns the value read by <see cref="ReadValue"/>, or <paramref name="defaultValue"/>. Called
    /// by compiled constructors for members whose values aren't typed.
    /// </summary>
    public static object? GetValueOrDefault(BoundMember<TOwner> member, object? defaultValue)
        => member.HasValue ? member.Value : defaultValue;

    /// <summary>
    /// Whether reading a value added no errors to <paramref name="errors"/>: a value that caused
    /// an error isn't used.
    /// </summary>
    private protected static bool AddedNoErrors(ErrorTarget? errors, int errorCount)
        => errors is null || errors.Count == errorCount;
}

/// <summary>
/// A bound member whose value is a <typeparamref name="TValue"/> from the map to the member's
/// setter, so value types aren't boxed.
/// </summary>
internal abstract class BoundMember<TOwner, TValue>(ClassMapMember member) : BoundMember<TOwner>(member)
{
    // Created on first use: members passed to a constructor (such as get-only properties) are
    // never set.
    private ExcelMemberSetter<TOwner, TValue>? _setter;

    // The value read by TryReadValue.
    private TValue _value = default!;

    /// <summary>
    /// Reads the member's value from the current row.
    /// </summary>
    /// <returns>Whether there is a value: false if the column is missing and the member is optional,
    /// or reading the value recorded an error.</returns>
    private protected abstract bool TryRead(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors, out TValue value);

    private ExcelMemberSetter<TOwner, TValue> Setter => _setter ??= CompiledAccessors.GetSetter<TOwner, TValue>(Member.Member);

    public sealed override bool TryMap(ref TOwner instance, SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors)
    {
        if (!TryRead(sheet, rowIndex, reader, errors, out var value))
        {
            return false;
        }

        Setter(ref instance, value);
        return true;
    }

    private protected sealed override bool TryReadValue(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors)
        => TryRead(sheet, rowIndex, reader, errors, out _value);

    public sealed override object? Value => _value;

    public sealed override void SetValue(ref TOwner instance) => Setter(ref instance, _value);

    /// <summary>
    /// Returns the value read by <see cref="BoundMember{TOwner}.ReadValue"/>, or
    /// <paramref name="defaultValue"/>, without boxing. Called by compiled constructors (see
    /// <see cref="IBindableMap.GetArgumentMethod{TOwner}(Type)"/>).
    /// </summary>
    public static TValue GetValueOrDefault(BoundMember<TOwner, TValue> member, TValue defaultValue)
        => member.HasValue ? member._value : defaultValue;
}

/// <summary>
/// A <see cref="ValueMap{T}"/> bound to a member.
/// </summary>
internal sealed class BoundValueMember<TOwner, T>(ValueMap<T> map, ClassMapMember member) : BoundMember<TOwner, T>(member)
{
    // The member's column in the sheet (-1 if missing), found on first use (inside the caller's
    // error handling, as for unbound maps).
    private int _columnIndex = -1;
    private bool _resolved;

    private protected override bool TryRead(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors, out T value)
    {
        if (!_resolved)
        {
            _columnIndex = map.FindColumn(sheet);
            _resolved = true;
        }

        var errorCount = errors?.Count ?? 0;
        return map.TryGetValue(_columnIndex, sheet, rowIndex, reader, Member.Member, errors, out value) && AddedNoErrors(errors, errorCount);
    }
}

/// <summary>
/// A nested <see cref="ExcelClassMap{T}"/> bound to a member: a nested struct isn't boxed either.
/// </summary>
internal sealed class BoundClassMember<TOwner, [DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> map, ClassMapMember member) : BoundMember<TOwner, T>(member)
{
    private protected override bool TryRead(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors, out T value)
    {
        var errorCount = errors?.Count ?? 0;
        return map.TryGetValue(sheet, rowIndex, reader, Member.Member, errors, out value) && AddedNoErrors(errors, errorCount);
    }
}

/// <summary>
/// A bound member whose map reads its value as an object.
/// </summary>
internal abstract class BoundObjectMember<TOwner>(ClassMapMember member) : BoundMember<TOwner, object?>(member)
{
    private protected sealed override bool TryRead(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors, out object? value)
    {
        var errorCount = errors?.Count ?? 0;
        if (TryGetValue(sheet, rowIndex, reader, errors, out value) && AddedNoErrors(errors, errorCount))
        {
            return true;
        }

        value = null;
        return false;
    }

    // Reads the value from the member's map.
    private protected abstract bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors, out object? value);
}

/// <summary>
/// Any other map bound to a member, mapped through <see cref="IMap.TryGetValue"/>: for example
/// indexer maps, row numbers and custom maps.
/// </summary>
internal sealed class BoundMapMember<TOwner>(ClassMapMember member) : BoundObjectMember<TOwner>(member)
{
    private protected override bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors, out object? value)
        => Member.Map.TryGetValue(sheet, rowIndex, reader, Member.Member, errors, out value);
}

/// <summary>
/// An <see cref="ICellsMap"/> bound to a member.
/// </summary>
internal sealed class BoundCellsMember<TOwner>(ICellsMap map, ClassMapMember member) : BoundObjectMember<TOwner>(member)
{
    private ICellsReader? _cellsReader;
    private bool _resolved;

    private protected override bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors, out object? value)
    {
        // Found on first use, inside the caller's error handling. Maps are frozen once rows are
        // read, so the reader stays right.
        if (!_resolved)
        {
            _cellsReader = map.GetCellsReader(sheet);
            _resolved = true;
        }

        return map.TryGetValue(_cellsReader, sheet, rowIndex, reader, Member.Member, errors, out value);
    }
}
