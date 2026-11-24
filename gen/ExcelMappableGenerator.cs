using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

namespace ExcelMapper.SourceGeneration;

/// <summary>
/// Writes the registrations of each <c>ExcelMappingContext</c> with <c>[ExcelMappable(typeof(T))]</c>:
/// <c>RegisterTypes</c>, registering each <c>T</c> as a row type and every type auto-mapping meets in
/// its members (see <see cref="TypeWalker"/>), a property for each row type's class map, and a static
/// <c>Default</c> instance.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class ExcelMappableGenerator : IIncrementalGenerator
{
    private const string AttributeName = "ExcelMapper.ExcelMappableAttribute";
    private const string ContextName = "ExcelMapper.ExcelMappingContext";

    /// <inheritdoc/>
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var contexts = context.SyntaxProvider.ForAttributeWithMetadataName(
            AttributeName,
            static (node, _) => node is ClassDeclarationSyntax,
            static (attributeContext, cancellationToken) => Parse((INamedTypeSymbol)attributeContext.TargetSymbol, attributeContext.SemanticModel.Compilation, cancellationToken));

        // A context declared in several partial parts, with attributes on more than one, is found
        // once for each part: write it once.
        context.RegisterSourceOutput(contexts.Collect(), static (sourceContext, models) =>
        {
            foreach (var model in models.GroupBy(m => m.FullName).Select(g => g.First()))
            {
                Emit(sourceContext, model);
            }
        });
    }

    private static ContextModel Parse(INamedTypeSymbol symbol, Compilation compilation, CancellationToken cancellationToken)
    {
        var diagnostics = ImmutableArray.CreateBuilder<DiagnosticModel>();
        var location = symbol.Locations.FirstOrDefault();
        var displayName = symbol.ToDisplayString();

        var derives = false;
        for (var baseType = symbol.BaseType; baseType is not null; baseType = baseType.BaseType)
        {
            if (baseType.ToDisplayString() == ContextName)
            {
                derives = true;
                break;
            }
        }

        if (!derives)
        {
            diagnostics.Add(DiagnosticModel.Create(Descriptors.ContextNotDerived, location, displayName));
        }

        var containingTypes = new List<string>();
        for (var containing = symbol.ContainingType; containing is not null; containing = containing.ContainingType)
        {
            containingTypes.Insert(0, Declaration(containing));
        }

        if (!IsPartial(symbol) || symbol.ContainingType is { } outer && !AllPartial(outer))
        {
            diagnostics.Add(DiagnosticModel.Create(Descriptors.ContextNotPartial, location, displayName));
        }

        var declaresRegisterTypes = symbol.GetMembers("RegisterTypes").Any(m => !m.IsImplicitlyDeclared);
        if (declaresRegisterTypes)
        {
            diagnostics.Add(DiagnosticModel.Create(Descriptors.RegisterTypesDeclared, location, displayName));
        }

        var walker = new TypeWalker(compilation);
        var properties = ImmutableArray.CreateBuilder<PropertyModel>();
        var memberNames = new HashSet<string>(symbol.MemberNames) { symbol.Name };
        foreach (var attribute in symbol.GetAttributes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (attribute.AttributeClass?.ToDisplayString() != AttributeName ||
                attribute.ConstructorArguments.Length != 1 ||
                attribute.ConstructorArguments[0].Value is not ITypeSymbol rowType)
            {
                continue;
            }

            if (rowType is INamedTypeSymbol { IsUnboundGenericType: true })
            {
                var attributeLocation = attribute.ApplicationSyntaxReference?.GetSyntax(cancellationToken).GetLocation() ?? location;
                diagnostics.Add(DiagnosticModel.Create(Descriptors.OpenGenericType, attributeLocation, rowType.ToDisplayString()));
                continue;
            }

            walker.AddRow(rowType);
            var rowLocation = attribute.ApplicationSyntaxReference?.GetSyntax(cancellationToken).GetLocation() ?? location;
            foreach (var (descriptor, arguments) in walker.TakeProblems())
            {
                diagnostics.Add(DiagnosticModel.Create(descriptor, rowLocation, arguments));
            }

            // A property named after the type, unless the context already has a member with that name.
            var propertyName = PropertyName(rowType);
            if (memberNames.Add(propertyName))
            {
                properties.Add(new PropertyModel(propertyName, TypeWalker.Name(rowType), rowType.ToDisplayString()));
            }
        }

        var generateDefault = derives && !symbol.IsAbstract && !symbol.IsStatic &&
            !symbol.MemberNames.Contains("Default") &&
            symbol.InstanceConstructors.Any(c => c.Parameters.Length == 0 && c.DeclaredAccessibility != Accessibility.Private || c.IsImplicitlyDeclared);

        return new ContextModel(
            symbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
            symbol.ContainingNamespace is { IsGlobalNamespace: false } ns ? ns.ToDisplayString() : null,
            new EquatableArray<string>(containingTypes.ToImmutableArray()),
            symbol.Name,
            generateDefault,
            derives && !declaresRegisterTypes,
            new EquatableArray<PropertyModel>(properties.ToImmutable()),
            new EquatableArray<string>(walker.Registrations.ToImmutableArray()),
            new EquatableArray<DiagnosticModel>(diagnostics.ToImmutable()));
    }

    private static bool IsPartial(INamedTypeSymbol type)
        => type.DeclaringSyntaxReferences.All(r => r.GetSyntax() is TypeDeclarationSyntax declaration && declaration.Modifiers.Any(SyntaxKind.PartialKeyword));

    private static bool AllPartial(INamedTypeSymbol type)
    {
        for (var current = type; current is not null; current = current.ContainingType)
        {
            if (!IsPartial(current))
            {
                return false;
            }
        }

        return true;
    }

    // The declaration of a containing type, to declare the context inside it.
    private static string Declaration(INamedTypeSymbol type)
    {
        var keyword = type switch
        {
            { IsRecord: true, TypeKind: TypeKind.Struct } => "record struct",
            { IsRecord: true } => "record",
            { TypeKind: TypeKind.Struct } => "struct",
            { TypeKind: TypeKind.Interface } => "interface",
            _ => "class",
        };

        var typeParameters = type.TypeParameters.Length > 0 ? $"<{string.Join(", ", type.TypeParameters.Select(p => p.Name))}>" : "";
        return $"partial {keyword} {type.Name}{typeParameters}";
    }

    // A name for the property of a row type's class map: the type's name, with its type arguments'
    // names for generic types, such as DictionaryStringInt32.
    private static string PropertyName(ITypeSymbol type) => type switch
    {
        IArrayTypeSymbol array => PropertyName(array.ElementType) + "Array",
        INamedTypeSymbol { IsGenericType: true } named => named.Name + string.Concat(named.TypeArguments.Select(PropertyName)),
        _ => type.Name,
    };

    private static void Emit(SourceProductionContext context, ContextModel model)
    {
        foreach (var diagnostic in model.Diagnostics)
        {
            context.ReportDiagnostic(diagnostic.ToDiagnostic());
        }

        if (model.Diagnostics.Items.Any(d => d.Id is "EXM001" or "EXM002"))
        {
            return;
        }

        const string GeneratedCode = "[global::System.CodeDom.Compiler.GeneratedCodeAttribute(\"ExcelMapper.SourceGeneration\", \"4.0.0.0\")]";
        var builder = new StringBuilder();
        builder.AppendLine("// <auto-generated/>");
        builder.AppendLine("#nullable enable");
        builder.AppendLine();
        var indent = "";
        if (model.Namespace is not null)
        {
            builder.AppendLine($"namespace {model.Namespace}");
            builder.AppendLine("{");
            indent += "    ";
        }

        foreach (var containingType in model.ContainingTypes)
        {
            builder.AppendLine($"{indent}{containingType}");
            builder.AppendLine($"{indent}{{");
            indent += "    ";
        }

        builder.AppendLine($"{indent}partial class {model.Name}");
        builder.AppendLine($"{indent}{{");
        var member = indent + "    ";
        var first = true;
        void Separate()
        {
            if (!first)
            {
                builder.AppendLine();
            }

            first = false;
        }

        if (model.GenerateDefault)
        {
            Separate();
            builder.AppendLine($"{member}/// <summary>Gets the default instance of the context.</summary>");
            builder.AppendLine($"{member}{GeneratedCode}");
            builder.AppendLine($"{member}public static {model.Name} Default {{ get; }} = new();");
        }

        foreach (var property in model.Properties)
        {
            Separate();
            builder.AppendLine($"{member}/// <summary>Gets the class map rows are mapped to <c>{Escape(property.DisplayName)}</c> with.</summary>");
            builder.AppendLine($"{member}{GeneratedCode}");
            builder.AppendLine($"{member}public global::ExcelMapper.ExcelClassMap<{property.TypeName}> {property.Name} => GetClassMap<{property.TypeName}>();");
        }

        if (model.GenerateRegisterTypes)
        {
            Separate();
            builder.AppendLine($"{member}/// <inheritdoc/>");
            builder.AppendLine($"{member}{GeneratedCode}");
            builder.AppendLine($"{member}protected override void RegisterTypes(global::ExcelMapper.ExcelTypeRegistrar types)");
            builder.AppendLine($"{member}{{");
            foreach (var registration in model.Registrations)
            {
                builder.AppendLine($"{member}    types.{registration};");
            }

            builder.AppendLine($"{member}}}");
        }

        builder.AppendLine($"{indent}}}");
        while (indent.Length > 0)
        {
            indent = indent.Substring(4);
            builder.AppendLine($"{indent}}}");
        }

        var hintName = model.FullName.Replace("global::", "").Replace('<', '[').Replace('>', ']') + ".g.cs";
        context.AddSource(hintName, SourceText.From(builder.ToString(), Encoding.UTF8));
    }

    private static string Escape(string text) => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
