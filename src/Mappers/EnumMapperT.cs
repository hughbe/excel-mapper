using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Mappers;

/// <summary>
/// An <see cref="EnumMapper"/> for a known enum type, which converts without boxing (see
/// <see cref="ICellConverter{T}"/>), for members of the enum type and of its nullable type.
/// </summary>
internal sealed class EnumMapper<TEnum> : EnumMapper, ICellConverter<TEnum>, ICellConverter<TEnum?> where TEnum : struct, Enum
{
    public EnumMapper(bool ignoreCase) : base(typeof(TEnum), ignoreCase)
    {
    }

    CellConversion ICellConverter<TEnum>.Convert(ReadCellResult cell, out TEnum value)
        => Enum.TryParse(cell.GetString(), IgnoreCase, out value) ? CellConversion.Success : CellConversion.Invalid;

    CellConversion ICellConverter<TEnum?>.Convert(ReadCellResult cell, out TEnum? value)
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
    public static EnumMapper Create(Type enumType, bool ignoreCase)
    {
        if (!enumType.IsEnum)
        {
            // Let EnumMapper report the error.
            return new EnumMapper(enumType, ignoreCase);
        }

        // Pass the argument as an array: on netstandard2.0, a bool binds to CreateInstance(Type, bool nonPublic).
        return (EnumMapper)Activator.CreateInstance(typeof(EnumMapper<>).MakeGenericType(enumType), [ignoreCase])!;
    }
}
