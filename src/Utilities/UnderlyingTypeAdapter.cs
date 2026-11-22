using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;

namespace ExcelMapper.Utilities;

/// <summary>
/// For a member of a nullable value type, adapts a converter or fallback of the underlying type
/// given by an attribute, such as an <see cref="ICellConverter{T}"/> of <see cref="int"/> on an
/// <c>int?</c> member. <see cref="TypeRule{T}.NullableAdapter"/> creates one per nullable type,
/// so adapting is typed rather than done by reflection for every member.
/// </summary>
/// <typeparam name="T">The member's type.</typeparam>
internal abstract class UnderlyingTypeAdapter<T>
{
    /// <summary>
    /// Adapts a converter of the underlying type, or returns null if it isn't one.
    /// </summary>
    public abstract ICellConverter<T>? AdaptConverter(object converter);

    /// <summary>
    /// Adapts a fallback of the underlying type, or returns null if it isn't one.
    /// </summary>
    public abstract ICellFallback<T>? AdaptFallback(object fallback);
}

/// <summary>
/// Adapts converters and fallbacks of <typeparamref name="T"/> for <c>T?</c> members.
/// </summary>
internal sealed class NullableAdapter<T> : UnderlyingTypeAdapter<T?> where T : struct
{
    public override ICellConverter<T?>? AdaptConverter(object converter)
        => converter is ICellConverter<T> typedConverter ? new NullableCellConverter<T>(typedConverter) : null;

    public override ICellFallback<T?>? AdaptFallback(object fallback)
        => fallback is ICellFallback<T> typedFallback ? new NullableCellFallback<T>(typedFallback) : null;
}
