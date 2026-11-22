using Microsoft.CodeAnalysis;

namespace ExcelMapper.SourceGeneration;

/// <summary>
/// The errors the generator reports.
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

    public static DiagnosticDescriptor Get(string id) => id switch
    {
        "EXM001" => ContextNotPartial,
        "EXM002" => ContextNotDerived,
        "EXM003" => RegisterTypesDeclared,
        _ => OpenGenericType,
    };
}
