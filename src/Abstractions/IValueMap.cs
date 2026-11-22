using System.Reflection;

namespace ExcelMapper.Abstractions;

/// <summary>
/// A map of a single value read from one cell (<see cref="ValueMap{T}"/>), which binds to a sheet
/// as a <see cref="BoundValueMember{T}"/>.
/// </summary>
internal interface IValueMap : IColumnsMap
{
    /// <summary>
    /// Binds the map to a member, for reading it from the rows of a sheet.
    /// </summary>
    BoundMember Bind(ClassMapMember member);

    /// <summary>
    /// Gets a static method <c>(TBoundMember member, TParameter defaultValue) → TParameter</c> that
    /// returns the value a bound member of this map read for a constructor parameter, without
    /// boxing. Null if the parameter's type isn't the map's type.
    /// </summary>
    MethodInfo? GetArgumentMethod(Type parameterType);
}
