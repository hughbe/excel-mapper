using System.ComponentModel;

namespace ExcelMapper;

/// <summary>
/// Calls a constructor of <typeparamref name="TOwner"/> with the values read for a row. The ExcelMapper
/// source generator writes one for each public constructor of a type an
/// <see cref="ExcelMappingContext"/> maps (see <see cref="ExcelTypeRegistrar.Constructor{TOwner}"/>),
/// so that native AOT apps create records and other types with constructor parameters without
/// reflection.
/// </summary>
/// <typeparam name="TOwner">The type created.</typeparam>
/// <param name="arguments">The constructor's arguments.</param>
/// <returns>The instance created.</returns>
[EditorBrowsable(EditorBrowsableState.Never)]
public delegate TOwner ExcelConstructor<TOwner>(ExcelConstructorArguments<TOwner> arguments);
