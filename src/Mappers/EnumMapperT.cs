using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Mappers;

/// <summary>
/// Maps an enum's names and numbers. Whether names are matched ignoring case comes from the cell's
/// <see cref="CellContext.Options"/> (<c>WithIgnoreCase</c>), so the mapper holds no state.
/// </summary>
internal sealed class EnumMapper<TEnum> : ICellMapper<TEnum>, ICellMapper<TEnum?> where TEnum : struct, Enum
{
    CellConversion ICellMapper<TEnum>.Convert(ReadCellResult cell, in CellContext context, out TEnum value, ref object? error)
        => Enum.TryParse(cell.GetString(), context.Options.IgnoreCase, out value) ? CellConversion.Success : CellConversion.Invalid;

    CellConversion ICellMapper<TEnum?>.Convert(ReadCellResult cell, in CellContext context, out TEnum? value, ref object? error)
    {
        if (Enum.TryParse(cell.GetString(), context.Options.IgnoreCase, out TEnum result))
        {
            value = result;
            return CellConversion.Success;
        }

        value = null;
        return CellConversion.Invalid;
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        var stringValue = cell.GetString();
        var ignoreCase = context.Options.IgnoreCase;

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
    public static object Create(Type enumType)
        => GenericTypes.Create(typeof(EnumMapper<>), enumType);
}
