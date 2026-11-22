using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Mappers;

internal sealed class EnumMapper<TEnum>(bool ignoreCase) : ICellMapper<TEnum>, ICellMapper<TEnum?> where TEnum : struct, Enum
{
    public bool IgnoreCase { get; } = ignoreCase;

    CellConversion ICellMapper<TEnum>.Convert(ReadCellResult cell, in CellContext context, out TEnum value, ref object? error)
        => Enum.TryParse(cell.GetString(), IgnoreCase, out value) ? CellConversion.Success : CellConversion.Invalid;

    CellConversion ICellMapper<TEnum?>.Convert(ReadCellResult cell, in CellContext context, out TEnum? value, ref object? error)
    {
        if (Enum.TryParse(cell.GetString(), IgnoreCase, out TEnum result))
        {
            value = result;
            return CellConversion.Success;
        }

        value = null;
        return CellConversion.Invalid;
    }

    public Exception CreateException(ReadCellResult cell, object? error)
    {
        var stringValue = cell.GetString();
        var ignoreCase = IgnoreCase;

        // Numbers can be out of range (an OverflowException): parse to find out.
        return stringValue is null || stringValue.Length == 0 || char.IsDigit(stringValue.TrimStart()[0]) || stringValue.TrimStart()[0] is '-' or '+'
            ? MapperUtilities.InvalidFromParse(() => Enum.Parse(typeof(TEnum), stringValue!, ignoreCase))
            : new ArgumentException($"The value \"{stringValue}\" isn't a value of {typeof(TEnum).Name}.");
    }
}

internal static class EnumMappers
{
    /// <summary>
    /// Creates the mapper for an enum type, which maps the type and its nullable version.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
#if NET5_0_OR_GREATER
    // The mapper type is cached, so the trimmer can't see it's created here.
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(EnumMapper<>))]
#endif
    public static object Create(Type enumType, bool ignoreCase)
    {
        // Pass the argument as an array: on netstandard2.0, a bool binds to CreateInstance(Type, bool nonPublic).
        var mapperType = s_mapperTypes.GetOrAdd(enumType, static t => typeof(EnumMapper<>).MakeGenericType(t));
        return Activator.CreateInstance(mapperType, [ignoreCase])!;
    }

    // EnumMapper<TEnum> for each enum type, so that it's created once.
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<Type, Type> s_mapperTypes = new();
}
