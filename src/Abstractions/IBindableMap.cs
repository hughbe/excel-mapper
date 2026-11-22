using System.Reflection;

namespace ExcelMapper.Abstractions;

/// <summary>
/// A map whose value has a type known when it's bound to a member of a class map: a
/// <see cref="ValueMap{T}"/> or a nested <see cref="ExcelClassMap{T}"/>. Its value stays typed from
/// the map to the member (see <see cref="BoundMember{TOwner, TValue}"/>), so value types aren't boxed.
/// </summary>
internal interface IBindableMap : IMap
{
    /// <summary>
    /// Binds the map to a member of a <typeparamref name="TOwner"/>, for reading it from the rows of a sheet.
    /// </summary>
    BoundMember<TOwner> Bind<TOwner>(ClassMapMember member);

    /// <summary>
    /// Gets a static method <c>(TBoundMember member, TParameter defaultValue) → TParameter</c> that
    /// returns the value a member of a <typeparamref name="TOwner"/> bound to this map read for a
    /// constructor parameter, without boxing. Null if the parameter's type isn't the map's type.
    /// </summary>
    MethodInfo? GetArgumentMethod<TOwner>(Type parameterType);
}
