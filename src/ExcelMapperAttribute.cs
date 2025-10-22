namespace ExcelMapper;

/// <summary>
/// Specifies a mapper to use for a property or field when reading or writing Excel files.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public sealed class ExcelMapperAttribute : Attribute
{
    /// <summary>
    /// Gets the type of the <see cref="ICellMapper"/>.
    /// </summary>
    public Type Type { get; init; }

    /// <summary>
    /// The constructor arguments for the <see cref="ICellMapper"/>.
    /// </summary>
    public object?[]? ConstructorArguments { get; set; }

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelMapperAttribute"/> with the specified transformer.
    /// </summary>
    /// <param name="mapperType">The type of the <see cref="ICellMapper"/>.</param>
    public ExcelMapperAttribute(Type mapperType)
    {
        ArgumentNullException.ThrowIfNull(mapperType);
        if (mapperType.IsAbstract || mapperType.IsInterface)
        {
            throw new ArgumentException("Mapper type cannot be abstract or an interface", nameof(mapperType));
        }
        if (!mapperType.ImplementsInterface(typeof(ICellMapper)))
        {
            throw new ArgumentException($"Mapper type must implement {nameof(ICellMapper)}", nameof(mapperType));
        }

        Type = mapperType;
    }
}

