using System.Runtime.CompilerServices;

namespace ExcelMapper;

internal static class ThrowHelpers
{
    public static void ThrowIfNull(object value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(value, paramName);
#else
        if (value is null)
        {
            throw new ArgumentNullException(paramName);
        }
#endif
    }

    public static void ThrowIfNullOrEmpty(string? value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
#if NET6_0_OR_GREATER
        ArgumentException.ThrowIfNullOrEmpty(value, paramName);
#else
        if (string.IsNullOrEmpty(value))
        {
            throw new ArgumentException("Value cannot be null or empty.", paramName);
        }
#endif
    }

    public static void ThrowIfNullOrWhiteSpace(string? value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
#if NET6_0_OR_GREATER
        ArgumentException.ThrowIfNullOrWhiteSpace(value, paramName);
#else
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("Value cannot be null or whitespace.", paramName);
        }
#endif
    }

    public static void ThrowIfNegative(int value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
#if NET6_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegative(value, paramName);
#else
        if (value < 0)
        {
            throw new ArgumentOutOfRangeException(paramName, "Value cannot be negative.");
        }
#endif
    }

    public static void ThrowIfNegativeOrZero(int value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
#if NET6_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(value, paramName);
#else
        if (value <= 0)
        {
            throw new ArgumentOutOfRangeException(paramName, "Value cannot be negative or zero.");
        }
#endif
    }

    public static void ThrowIfLessThanOrEqual(int value, int minValue, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
#if NET6_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(value, minValue, paramName);
#else
        if (value <= minValue)
        {
            throw new ArgumentOutOfRangeException(paramName, $"Value cannot be less than or equal to {minValue}.");
        }
#endif
    }

    public static void ThrowIfGreaterThan(int value, int maxValue, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
#if NET6_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfGreaterThan(value, maxValue, paramName);
#else
        if (value > maxValue)
        {
            throw new ArgumentOutOfRangeException(paramName, $"Value cannot be greater than {maxValue}.");
        }
#endif
    }

    public static void ThrowIfGreaterThanOrEqual(int value, int maxValue, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
#if NET6_0_OR_GREATER
    ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(value, maxValue, paramName);
#else
        if (value >= maxValue)
        {
            throw new ArgumentOutOfRangeException(paramName, $"Value cannot be greater than or equal to {maxValue}.");
        }
#endif
    }
}
