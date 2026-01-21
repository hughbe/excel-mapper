using System.Globalization;

namespace ExcelMapper.Utilities;

public static class StringComparerUtilities
{
    public static StringComparer FromComparison(StringComparison comparisonType)
    {
#if NETSTANDARD2_1 || NETCOREAPP2_0_OR_GREATER
        return StringComparer.FromComparison(comparisonType);
#else
        return comparisonType switch
        {
            StringComparison.CurrentCulture => StringComparer.CurrentCulture,
            StringComparison.CurrentCultureIgnoreCase => StringComparer.CurrentCultureIgnoreCase,
            StringComparison.InvariantCulture => StringComparer.InvariantCulture,
            StringComparison.InvariantCultureIgnoreCase => StringComparer.InvariantCultureIgnoreCase,
            StringComparison.Ordinal => StringComparer.Ordinal,
            StringComparison.OrdinalIgnoreCase => StringComparer.OrdinalIgnoreCase,
            _ => throw new ArgumentException("String comparison type not supported.", nameof(comparisonType)),
        };
#endif
    }
}