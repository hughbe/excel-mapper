namespace ExcelMapper;

/// <summary>
/// Sets the property or field to the number of the row being read instead of reading a column.
/// Row numbers are one-based, as Excel shows them, and count every row in the sheet including
/// the heading, so they match the rows in error messages.
/// </summary>
/// <remarks>
/// The member must have a numeric type such as <see cref="int"/> or <see cref="long"/>.
/// For a positional record, apply it to the property: <c>[property: ExcelRowNumber]</c>.
/// </remarks>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelRowNumberAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelRowNumberAttribute"/>.
    /// </summary>
    public ExcelRowNumberAttribute()
    {
    }
}
