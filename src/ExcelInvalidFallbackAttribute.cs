namespace ExcelMapper;

/// <summary>
/// Specifies the empty <see cref="IFallbackItem"/> to use for a property or field when reading or writing Excel files.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public sealed class ExcelInvalidFallbackAttribute : Attribute
{
    /// <summary>
    /// Gets the type of the <see cref="IFallbackItem"/>.
    /// </summary>
    public Type Type { get; }

    /// <summary>
    /// The constructor arguments for the <see cref="IFallbackItem"/>.
    /// </summary>
    public object?[]? ConstructorArguments { get; set; }

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelInvalidFallbackAttribute"/> with the specified transformer.
    /// </summary>
    /// <param name="fallbackType">The type of the <see cref="IFallbackItem"/>.</param>
    public ExcelInvalidFallbackAttribute(Type fallbackType)
    {
        ThrowHelpers.ThrowIfNull(fallbackType, nameof(fallbackType));
        if (fallbackType.IsAbstract || fallbackType.IsInterface)
        {
            throw new ArgumentException("Fallback type cannot be abstract or an interface", nameof(fallbackType));
        }
        if (!fallbackType.ImplementsInterface(typeof(IFallbackItem)))
        {
            throw new ArgumentException($"Fallback type must implement {nameof(IFallbackItem)}", nameof(fallbackType));
        }

        Type = fallbackType;
    }
}

