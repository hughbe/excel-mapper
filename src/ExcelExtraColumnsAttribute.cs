namespace ExcelMapper;

/// <summary>
/// Fills a dictionary property or field with every column in the heading that no other member
/// of the same class reads, keyed by column name. Use it to keep columns a class doesn't know
/// about, for example to pass them on or report them.
/// </summary>
/// <remarks>
/// The member must be a dictionary with string keys, such as
/// <see cref="Dictionary{TKey, TValue}"/> of <see cref="string"/> to <see cref="string"/> or
/// <see cref="object"/>, or <see cref="System.Dynamic.ExpandoObject"/>. The sheet must have a heading.
/// For a positional record, apply it to the property: <c>[property: ExcelExtraColumns]</c>.
/// </remarks>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelExtraColumnsAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelExtraColumnsAttribute"/>.
    /// </summary>
    public ExcelExtraColumnsAttribute()
    {
    }
}
