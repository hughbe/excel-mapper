namespace ExcelMapper.Abstractions;

/// <summary>
/// A map of a single value read from one cell (<see cref="ValueMap{T}"/>), which binds to a sheet
/// as a <see cref="BoundValueMember{TOwner, T}"/>.
/// </summary>
internal interface IValueMap : IColumnsMap, IBindableMap
{
}
