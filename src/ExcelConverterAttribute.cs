using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper;

/// <summary>
/// Converts the member's cells with an <see cref="ICellConverter{T}"/>, replacing the built-in
/// conversion for the member's type. For a collection, the converter converts each element.
/// </summary>
/// <remarks>
/// The converter type must have a public constructor that takes <see cref="ConstructorArguments"/>
/// (or no arguments), and implement <see cref="ICellConverter{T}"/> for the member's type, or for
/// <c>T</c> if the member's type is <c>T?</c>. Several converters are tried in order: a converter
/// that returns <see cref="CellConversionResult{T}.NotHandled"/> lets the next one try.
/// </remarks>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public sealed class ExcelConverterAttribute : Attribute
{
    /// <summary>
    /// Gets the type of the converter.
    /// </summary>
    [DynamicallyAccessedMembers(TrimmingConstants.ActivatedType)]
    public Type Type { get; }

    /// <summary>
    /// Gets or sets the arguments passed to the converter's constructor.
    /// </summary>
    public object?[]? ConstructorArguments { get; set; }

    /// <summary>
    /// Converts the member's cells with a converter of type <paramref name="converterType"/>.
    /// </summary>
    /// <param name="converterType">A type that implements <see cref="ICellConverter{T}"/>.</param>
    /// <exception cref="ArgumentException"><paramref name="converterType"/> is abstract, an
    /// interface, or doesn't implement <see cref="ICellConverter{T}"/>.</exception>
    public ExcelConverterAttribute([DynamicallyAccessedMembers(TrimmingConstants.ActivatedType)] Type converterType)
    {
        ThrowHelpers.ThrowIfNull(converterType, nameof(converterType));
        if (converterType.IsAbstract || converterType.IsInterface)
        {
            throw new ArgumentException($"Converter type \"{converterType}\" can't be abstract or an interface.", nameof(converterType));
        }

        if (!converterType.ImplementsGenericInterface(typeof(ICellConverter<>), out _))
        {
            throw new ArgumentException($"Converter type \"{converterType}\" must implement ICellConverter<T>.", nameof(converterType));
        }

        Type = converterType;
    }
}
