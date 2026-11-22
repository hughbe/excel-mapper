using System.ComponentModel;

namespace ExcelMapper;

/// <summary>
/// Sets a property or field of <paramref name="owner"/>, which is passed by reference so that a
/// struct is set in place. The ExcelMapper source generator writes one for each member an
/// <see cref="ExcelMappingContext"/> maps (see <see cref="ExcelTypeRegistrar.Setter{TOwner, TValue}"/>),
/// so that native AOT apps set members without reflection.
/// </summary>
/// <typeparam name="TOwner">The type whose member is set.</typeparam>
/// <typeparam name="TValue">The member's type.</typeparam>
/// <param name="owner">The instance whose member is set.</param>
/// <param name="value">The value to set.</param>
[EditorBrowsable(EditorBrowsableState.Never)]
public delegate void ExcelMemberSetter<TOwner, in TValue>(ref TOwner owner, TValue value);
