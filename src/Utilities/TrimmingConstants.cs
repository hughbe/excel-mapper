using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Utilities;

/// <summary>
/// Trimming annotations and messages shared across the library.
/// </summary>
internal static class TrimmingConstants
{
    /// <summary>
    /// The members of a type given in an attribute (a mapper, matcher, fallback or transformer)
    /// that the library uses: it checks the interfaces and calls a public constructor.
    /// </summary>
    public const DynamicallyAccessedMemberTypes ActivatedType = DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.Interfaces;

    /// <summary>
    /// The members of a class map's type that are used to create it and set its members.
    /// </summary>
    public const DynamicallyAccessedMemberTypes MappedType =
        DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields;

    /// <summary>
    /// The members of a collection or dictionary type that its builder uses to create it and add items.
    /// </summary>
    public const DynamicallyAccessedMemberTypes CollectionType =
        DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicMethods | DynamicallyAccessedMemberTypes.Interfaces;

    public const string AutoMappingMessage =
        "Mapping discovers and sets properties, fields and constructors through reflection, including those of nested types, " +
        "collections and dictionaries, which the trimmer can't see. Make sure the mapped types and their members are preserved.";

    public const string AutoMappingDynamicCodeMessage =
        "Mapping creates generic types and methods for the mapped member types at run time, which native AOT can't guarantee.";
}
