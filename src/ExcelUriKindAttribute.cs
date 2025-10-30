namespace ExcelMapper;

/// <summary>
/// An attribute used to specify the kind of URI to use when mapping cell values to <see cref="Uri"/> objects.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelUriAttribute : Attribute
{
    /// <summary>
    /// Gets the URI kind.
    /// </summary>
    public UriKind UriKind { get; }

    /// <summary>
    /// Constructs the attribute with the specified URI kind.
    /// </summary>
    /// <param name="uriKind">The kind of URI to use when mapping cell values to <see cref="Uri"/> objects. Must be a valid <see cref="System.UriKind"/> value (Absolute, Relative, or RelativeOrAbsolute).</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="uriKind"/> is not a valid <see cref="System.UriKind"/> value.</exception>
    public ExcelUriAttribute(UriKind uriKind)
    {
        EnumUtilities.ValidateIsDefined(uriKind);
        UriKind = uriKind;
    }
}
