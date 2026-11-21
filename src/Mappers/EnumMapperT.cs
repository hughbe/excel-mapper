using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Mappers;

/// <summary>
/// An <see cref="EnumMapper"/> for a known enum type, which converts without boxing (see
/// <see cref="ITypedCellMapper{T}"/>), for members of the enum type and of its nullable type.
/// </summary>
internal sealed class EnumMapper<TEnum> : EnumMapper, ITypedCellMapper<TEnum>, ITypedCellMapper<TEnum?> where TEnum : struct, Enum
{
    public EnumMapper(bool ignoreCase) : base(typeof(TEnum), ignoreCase)
    {
    }

    CellConversion ITypedCellMapper<TEnum>.Convert(ReadCellResult cell, out TEnum value)
        => Enum.TryParse(cell.GetString(), IgnoreCase, out value) ? CellConversion.Success : CellConversion.Invalid;

    CellConversion ITypedCellMapper<TEnum?>.Convert(ReadCellResult cell, out TEnum? value)
    {
        if (Enum.TryParse(cell.GetString(), IgnoreCase, out TEnum result))
        {
            value = result;
            return CellConversion.Success;
        }

        value = null;
        return CellConversion.Invalid;
    }
}

internal static class EnumMappers
{
    /// <summary>
    /// Creates an <see cref="EnumMapper"/> for an enum type, which converts without boxing.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
#if NET5_0_OR_GREATER
    // The mapper type is cached, so the trimmer can't see it's created here.
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(EnumMapper<>))]
#endif
    public static EnumMapper Create(Type enumType, bool ignoreCase)
    {
        if (!enumType.IsEnum)
        {
            // Let EnumMapper report the error.
            return new EnumMapper(enumType, ignoreCase);
        }

        // Pass the argument as an array: on netstandard2.0, a bool binds to CreateInstance(Type, bool nonPublic).
        var mapperType = s_mapperTypes.GetOrAdd(enumType, static t => typeof(EnumMapper<>).MakeGenericType(t));
        return (EnumMapper)Activator.CreateInstance(mapperType, [ignoreCase])!;
    }

    // EnumMapper<TEnum> for each enum type, so that it's created once.
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<Type, Type> s_mapperTypes = new();
}
