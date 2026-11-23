using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

internal static class SeparatorUtilities
{
    public static void ValidateSeparators(string[] separators, [CallerArgumentExpression(nameof(separators))] string? paramName = null)
    {
        ThrowHelpers.ThrowIfNull(separators, paramName);
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

    // RemoveEmptyEntries and TrimEntries, which the netstandard builds don't name.
    private const StringSplitOptions AllSplitOptions = (StringSplitOptions)3;

    public static void ValidateOptions(StringSplitOptions options, [CallerArgumentExpression(nameof(options))] string? paramName = null)
    {
        if ((options & ~AllSplitOptions) != 0)
        {
            throw new ArgumentOutOfRangeException(paramName, options, "Split options can only remove empty entries and trim entries.");
        }
    }

    public static void ValidateSeparators(char[] separators, [CallerArgumentExpression(nameof(separators))] string? paramName = null)
    {
        ThrowHelpers.ThrowIfNull(separators, paramName);
        if (separators.Length == 0)
        {
            throw new ArgumentException("Separators cannot be empty.", paramName);
        }
    }
}
