using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

internal static class SeparatorUtilities
{
    public static void ValidateSeparators(string[] separators, [CallerArgumentExpression(nameof(separators))] string? paramName = null)
    {
        ArgumentNullException.ThrowIfNull(separators, paramName);
        if (separators.Length == 0)
        {
            throw new ArgumentException("Separators cannot be empty.", paramName);
        }

        foreach (var separator in separators)
        {
            if (string.IsNullOrEmpty(separator))
            {
                throw new ArgumentException("Separators cannot contain null or empty values.", paramName);
            }
        }
    }

    public static void ValidateSeparators(char[] separators, [CallerArgumentExpression(nameof(separators))] string? paramName = null)
    {
        ArgumentNullException.ThrowIfNull(separators, paramName);
        if (separators.Length == 0)
        {
            throw new ArgumentException("Separators cannot be empty.", paramName);
        }
    }
}
