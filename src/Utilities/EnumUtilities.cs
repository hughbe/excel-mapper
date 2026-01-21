using System.Runtime.CompilerServices;

namespace ExcelMapper;

internal class EnumUtilities
{
    public static void ValidateIsDefined<TEnum>(TEnum value, [CallerArgumentExpression(nameof(value))] string? paramName = null) where TEnum : struct, Enum
    {
#if NET5_0_OR_GREATER
        if (!Enum.IsDefined(value))
#else
        if (!Enum.IsDefined(typeof(TEnum), value))
#endif
        {
            throw new ArgumentOutOfRangeException(paramName, value, $"Invalid value \"{value}\".");
        }
    }
}
