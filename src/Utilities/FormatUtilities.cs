using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

internal static class FormatUtilities
{
    public static void ValidateFormats(string[] formats, [CallerArgumentExpression(nameof(formats))] string? paramName = null)
    {
        ArgumentNullException.ThrowIfNull(formats, paramName);
        if (formats.Length == 0)
        {
            throw new ArgumentException("Formats cannot be empty.", paramName);
        }
        foreach (var format in formats)
        {
            if (string.IsNullOrEmpty(format))
            {
                throw new ArgumentException("Formats cannot contain null or empty values.", paramName);
            }
        }
    }
}