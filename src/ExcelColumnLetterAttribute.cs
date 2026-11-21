namespace ExcelMapper;

/// <summary>
/// Maps a property or field to the column with the given Excel letters, such as "A", "C" or "AB".
/// This is the same as <see cref="ExcelColumnIndexAttribute"/> with the column's zero-based index.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public class ExcelColumnLetterAttribute : ExcelColumnIndexAttribute
{
    /// <summary>
    /// Maps a property or field to the column with the given Excel letters.
    /// </summary>
    /// <param name="letter">The column's letters, such as "A", "C" or "AB". Case-insensitive.</param>
    public ExcelColumnLetterAttribute(string letter) : base(ExcelRange.ParseColumnLetters(letter, nameof(letter)))
    {
    }

    /// <summary>
    /// Gets the column's letters, such as "C". This always matches <see cref="ExcelColumnIndexAttribute.Index"/>.
    /// </summary>
    public string Letter => ExcelRange.GetColumnLetters(Index);
}
