using System.Diagnostics.CodeAnalysis;
namespace ExcelMapper;

/// <summary>
/// Specifies a transformer to use for a property or field when reading or writing Excel files.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public sealed class ExcelTransformerAttribute : Attribute
{
    /// <summary>
    /// Gets the type of the <see cref="Abstractions.ICellTransformer"/>.
    /// </summary>
    [DynamicallyAccessedMembers(TrimmingConstants.ActivatedType)]
    public Type Type { get; }

    /// <summary>
    /// The constructor arguments for the <see cref="Abstractions.ICellTransformer"/>.
    /// </summary>
    public object?[]? ConstructorArguments { get; set; }

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelTransformerAttribute"/> with the specified transformer.
    /// </summary>
    /// <param name="transformerType">The type of the <see cref="Abstractions.ICellTransformer"/>.</param>
    public ExcelTransformerAttribute([DynamicallyAccessedMembers(TrimmingConstants.ActivatedType)] Type transformerType)
    {
        ThrowHelpers.ThrowIfNull(transformerType, nameof(transformerType));
        if (transformerType.IsAbstract || transformerType.IsInterface)
        {
            throw new ArgumentException("Transformer type cannot be abstract or an interface", nameof(transformerType));
        }
        if (!transformerType.ImplementsInterface(typeof(ICellTransformer)) && !transformerType.ImplementsInterface(typeof(Abstractions.ICellTransformer)))
        {
            throw new ArgumentException($"Transformer type must implement {nameof(ICellTransformer)}", nameof(transformerType));
        }

        Type = transformerType;
    }
}

