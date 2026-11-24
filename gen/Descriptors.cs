using Microsoft.CodeAnalysis;

namespace ExcelMapper.SourceGeneration;

/// <summary>
/// The errors and warnings the generator reports.
/// </summary>
internal static class Descriptors
{
    private const string Category = "ExcelMapper.SourceGeneration";

    public static DiagnosticDescriptor ContextNotPartial { get; } = new(
        "EXM001",
        "The mapping context must be partial",
        "The mapping context '{0}', and every type it's declared in, must be partial so that the source generator can add its registrations",
        Category,
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static DiagnosticDescriptor ContextNotDerived { get; } = new(
        "EXM002",
        "The mapping context must derive from ExcelMappingContext",
        "'{0}' has [ExcelMappable] but doesn't derive from ExcelMapper.ExcelMappingContext",
        Category,
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static DiagnosticDescriptor RegisterTypesDeclared { get; } = new(
        "EXM003",
        "RegisterTypes is written by the source generator",
        "The mapping context '{0}' declares RegisterTypes, which the source generator writes from its [ExcelMappable] attributes: remove it",
        Category,
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static DiagnosticDescriptor OpenGenericType { get; } = new(
        "EXM004",
        "Rows can't be mapped to an open generic type",
        "Rows can't be mapped to the open generic type '{0}': give its type arguments, such as List<int>",
        Category,
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    // Warnings, rather than errors: they mirror auto-mapping's rules (see TypeWalker), and a
    // difference shouldn't stop an app building. Each would make GetClassMap throw for the row type.
    public static DiagnosticDescriptor NothingToMap { get; } = new(
        "EXM005",
        "The type has nothing auto-mapping can set",
        "'{0}' has no members auto-mapping can set, so reading rows of '{1}' will throw: give it public setters, or a public constructor whose parameters are its properties",
        Category,
        DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public static DiagnosticDescriptor CannotCreate { get; } = new(
        "EXM006",
        "Auto-mapping can't create the type",
        "Auto-mapping can't create '{0}', so reading rows of '{1}' will throw: {2}",
        Category,
        DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public static DiagnosticDescriptor ContainsItself { get; } = new(
        "EXM007",
        "The type contains itself",
        "'{0}' contains itself through member '{1}', so reading rows of '{2}' will throw: ignore the member with [ExcelIgnore]",
        Category,
        DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public static DiagnosticDescriptor Get(string id) => id switch
    {
        "EXM001" => ContextNotPartial,
        "EXM002" => ContextNotDerived,
        "EXM003" => RegisterTypesDeclared,
        "EXM004" => OpenGenericType,
        "EXM005" => NothingToMap,
        "EXM006" => CannotCreate,
        _ => ContainsItself,
    };
}
