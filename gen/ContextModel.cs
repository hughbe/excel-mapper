using System.Collections.Immutable;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;

namespace ExcelMapper.SourceGeneration;

/// <summary>
/// What the generator writes for a context: where it's declared, the class map properties, the
/// registrations, and any errors. It holds only strings and values, so it compares by value.
/// </summary>
internal sealed record ContextModel(
    string FullName,
    string? Namespace,
    EquatableArray<string> ContainingTypes,
    string Name,
    bool GenerateDefault,
    bool GenerateRegisterTypes,
    EquatableArray<PropertyModel> Properties,
    EquatableArray<string> Registrations,
    EquatableArray<DiagnosticModel> Diagnostics);

/// <summary>
/// A property that gives the class map of a type rows are mapped to.
/// </summary>
internal sealed record PropertyModel(string Name, string TypeName, string DisplayName);

/// <summary>
/// An error or warning to report, with where it is and the arguments of its message.
/// </summary>
internal sealed record DiagnosticModel(string Id, string? FilePath, TextSpan Span, LinePositionSpan LineSpan, EquatableArray<string> Arguments)
{
    public static DiagnosticModel Create(DiagnosticDescriptor descriptor, Location? location, params string[] arguments)
        => new(descriptor.Id, location?.SourceTree?.FilePath, location?.SourceSpan ?? default, location?.GetLineSpan().Span ?? default, new EquatableArray<string>(arguments.ToImmutableArray()));

    public Diagnostic ToDiagnostic()
    {
        var location = FilePath is null ? Location.None : Location.Create(FilePath, Span, LineSpan);
        return Diagnostic.Create(Descriptors.Get(Id), location, Arguments.Items.ToArray<object>());
    }
}
