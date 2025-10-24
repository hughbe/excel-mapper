using System.Runtime.CompilerServices;

namespace ExcelMapper;

internal class EnumUtilities
{
    public static void ValidateIsDefined<TEnum>(TEnum value, [CallerArgumentExpression(nameof(value))] string? paramName = null) where TEnum : struct, Enum
    {
        if (!Enum.IsDefined(value))
        {
            throw new ArgumentOutOfRangeException(paramName, value, $"Invalid value \"{value}\".");
        }
    }
}
