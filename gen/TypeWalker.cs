using System.Collections.Generic;
using System.Linq;
using Microsoft.CodeAnalysis;

namespace ExcelMapper.SourceGeneration;

/// <summary>
/// Finds every type auto-mapping meets when it maps rows to a type, and the
/// <c>ExcelTypeRegistrar</c> call that registers each, in the order <c>AutoMapper.TryCreateMap</c>
/// tries them: a built-in conversion (including <c>IConvertible</c>, number and
/// <c>IParsable&lt;T&gt;</c> types), a dictionary, a collection, then an object mapped through its
/// public properties and fields.
/// </summary>
/// <remarks>
/// <para>Registering a type that isn't needed only adds code to the app; missing one makes auto-mapping
/// fail with native AOT. So where auto-mapping could go either way, both are registered.</para>
/// <para>This mirrors auto-mapping in the library, and the two must change together: <c>Visit</c>
/// follows <c>AutoMapper.TryCreateMap</c>, <c>HasConversion</c> and <c>s_builtInTypes</c> follow
/// <c>TypeRule&lt;T&gt;.FindDefaultMappers</c>, <c>CanConstructObject</c> follows
/// <c>AutoMapper.CanConstructObject</c>, <c>TryGetKeyValueTypes</c> follows
/// <c>DictionaryRule.TryGetKeyValueTypes</c>, <c>TryGetElementType</c> follows
/// <c>ReflectionUtilities.GetElementTypeOrEnumerableType</c>, and <c>GetMemberTypes</c> follows
/// <c>AutoMapper.TryAddMembers</c>. See "Auto-mapping and the source generator" in CLAUDE.md.</para>
/// </remarks>
internal sealed class TypeWalker
{
    // The types with a built-in conversion: TypeRule<T> creates their mappers without generic types
    // made at run time.
    private static readonly HashSet<string> s_builtInTypes =
    [
        "System.String", "System.Object", "System.Boolean", "System.Char", "System.SByte", "System.Byte",
        "System.Int16", "System.UInt16", "System.Int32", "System.UInt32", "System.Int64", "System.UInt64",
        "System.Single", "System.Double", "System.Decimal", "System.DateTime", "System.DateTimeOffset",
        "System.TimeSpan", "System.Guid", "System.Uri", "System.Version", "System.IConvertible", "System.DateOnly",
        "System.TimeOnly", "System.Half", "System.Int128", "System.UInt128", "System.IntPtr", "System.UIntPtr",
        "System.Numerics.BigInteger", "System.Numerics.Complex",
    ];

    private static readonly SymbolDisplayFormat s_typeFormat = SymbolDisplayFormat.FullyQualifiedFormat;

    private readonly Compilation _compilation;
    private readonly INamedTypeSymbol? _enumerableOfT;
    private readonly INamedTypeSymbol? _enumerable;
    private readonly INamedTypeSymbol? _dictionary;
    private readonly INamedTypeSymbol? _keyValuePair;
    private readonly INamedTypeSymbol? _convertible;
    private readonly INamedTypeSymbol? _parsable;
    private readonly INamedTypeSymbol? _numberBase;
    private readonly INamedTypeSymbol? _ignoreAttribute;
    private readonly bool _canRegisterParsable;

    private readonly HashSet<ITypeSymbol> _visited = new(SymbolEqualityComparer.Default);
    private readonly List<string> _registrations = [];
    private readonly HashSet<string> _registered = [];

    public TypeWalker(Compilation compilation)
    {
        _compilation = compilation;
        _enumerableOfT = compilation.GetSpecialType(SpecialType.System_Collections_Generic_IEnumerable_T);
        _enumerable = compilation.GetSpecialType(SpecialType.System_Collections_IEnumerable);
        _dictionary = compilation.GetTypeByMetadataName("System.Collections.IDictionary");
        _keyValuePair = compilation.GetTypeByMetadataName("System.Collections.Generic.KeyValuePair`2");
        _convertible = compilation.GetTypeByMetadataName("System.IConvertible");
        _parsable = compilation.GetTypeByMetadataName("System.IParsable`1");
        _numberBase = compilation.GetTypeByMetadataName("System.Numerics.INumberBase`1");
        _ignoreAttribute = compilation.GetTypeByMetadataName("ExcelMapper.ExcelIgnoreAttribute");

        // The library's netstandard builds don't have IParsable<T> registrations.
        _canRegisterParsable = compilation.GetTypeByMetadataName("ExcelMapper.ExcelTypeRegistrar")?.GetMembers("Parsable").Length > 0;
    }

    public IReadOnlyList<string> Registrations => _registrations;

    /// <summary>
    /// Registers a type rows are mapped to, and every type it uses.
    /// </summary>
    public void AddRow(ITypeSymbol type)
    {
        Add($"Row<{Name(type)}>()");

        // Row registers the type as Type does.
        _registered.Add($"Type<{Name(type)}>()");
        Visit(type);
    }

    public static string Name(ITypeSymbol type) => type.WithNullableAnnotation(NullableAnnotation.NotAnnotated).ToDisplayString(s_typeFormat);

    private void Add(string registration)
    {
        if (_registered.Add(registration))
        {
            _registrations.Add(registration);
        }
    }

    private void Visit(ITypeSymbol type)
    {
        if (type.TypeKind is TypeKind.Error or TypeKind.Pointer or TypeKind.FunctionPointer or TypeKind.Dynamic || !_visited.Add(type))
        {
            return;
        }

        var name = Name(type);
        if (type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable)
        {
            var underlyingType = nullable.TypeArguments[0];
            Add($"Nullable<{Name(underlyingType)}>()");
            Visit(underlyingType);
            return;
        }

        if (type.TypeKind == TypeKind.Enum)
        {
            Add($"Enum<{name}>()");
            return;
        }

        if (type is IArrayTypeSymbol array)
        {
            // Multidimensional arrays are only mapped by index with Map(...).
            if (array.Rank == 1)
            {
                Add($"Array<{Name(array.ElementType)}>()");
                Visit(array.ElementType);
            }

            return;
        }

        if (s_builtInTypes.Contains(MetadataName(type)))
        {
            Add($"Type<{name}>()");
            return;
        }

        // The library's netstandard builds map IParsable<T> types through their members, as objects.
        if (HasConversion(type, out var isParsable) && (!isParsable || _canRegisterParsable))
        {
            Add(isParsable ? (type.IsValueType ? $"Parsable<{name}>()" : $"ParsableClass<{name}>()") : $"Type<{name}>()");
            return;
        }

        if (TryGetKeyValueTypes(type, out var keyType, out var valueType))
        {
            Add($"Dictionary<{name}, {Name(keyType)}, {Name(valueType)}>()");
            Visit(keyType);
            Visit(valueType);
            return;
        }

        if (TryGetElementType(type, out var elementType))
        {
            Add($"Collection<{name}, {Name(elementType)}>()");
            Visit(elementType);
            return;
        }

        Add($"Type<{name}>()");
        AddConstructors(type, name);
        foreach (var member in GetMappedMembers(type))
        {
            var memberType = MemberType(member);
            AddSetter(member, memberType, name);
            Visit(memberType);
        }
    }

    // Registers a setter for a member auto-mapping sets: a property with a public setter that
    // isn't init-only, or a field that isn't readonly. Others are set through a constructor, or by
    // reflection (see CompiledAccessors.GetSetter).
    private void AddSetter(ISymbol member, ITypeSymbol memberType, string ownerName)
    {
        var settable = member switch
        {
            IPropertySymbol { SetMethod: { DeclaredAccessibility: Accessibility.Public, IsInitOnly: false } } => true,
            IFieldSymbol { IsReadOnly: false } => true,
            _ => false,
        };
        if (!settable || memberType.IsRefLikeType)
        {
            return;
        }

        var valueName = Name(memberType);
        var identifier = Identifier(member.Name);
        Add($"Setter<{ownerName}, {valueName}>(\"{member.Name}\", static (ref {ownerName} owner, {valueName} value) => owner.{identifier} = value)");
    }

    // Registers the public constructors with parameters of a type auto-mapping creates, for records
    // and types with read-only members (see ConstructionPlan<T>.Create).
    private void AddConstructors(ITypeSymbol type, string name)
    {
        if (type is not INamedTypeSymbol { IsAbstract: false, TypeKind: TypeKind.Class or TypeKind.Struct } named)
        {
            return;
        }

        foreach (var constructor in named.InstanceConstructors)
        {
            if (constructor.DeclaredAccessibility != Accessibility.Public || constructor.Parameters.Length == 0 ||
                constructor.Parameters.Any(p => p.RefKind != RefKind.None || p.Type.IsRefLikeType))
            {
                continue;
            }

            var parameterTypes = string.Join(", ", constructor.Parameters.Select(p => $"typeof({Name(p.Type)})"));
            var arguments = string.Join(", ", constructor.Parameters.Select((p, i) => $"arguments.Get<{Name(p.Type)}>({i})"));
            Add($"Constructor<{name}>(new global::System.Type[] {{ {parameterTypes} }}, static arguments => new {name}({arguments}))");
        }
    }

    private static ITypeSymbol MemberType(ISymbol member) => member is IPropertySymbol property ? property.Type : ((IFieldSymbol)member).Type;

    // A member's name as an identifier, escaping keywords such as "class".
    private static string Identifier(string name)
        => Microsoft.CodeAnalysis.CSharp.SyntaxFacts.GetKeywordKind(name) != Microsoft.CodeAnalysis.CSharp.SyntaxKind.None ? "@" + name : name;

    // Whether auto-mapping converts the type's cells itself (see TypeRule<T>.FindDefaultMappers):
    // number and IParsable<T> types other than interfaces, which their Parse methods create, and
    // IConvertible types it can create.
    private bool HasConversion(ITypeSymbol type, out bool isParsable)
    {
        isParsable = false;
        if (type.TypeKind != TypeKind.Interface && ImplementsGeneric(type, _numberBase, out _))
        {
            return true;
        }

        if (CanConstructObject(type) && Implements(type, _convertible))
        {
            return true;
        }

        isParsable = type.TypeKind != TypeKind.Interface && ImplementsGeneric(type, _parsable, out _);
        return isParsable;
    }

    // A value type, or a class with a public parameterless constructor (see AutoMapper.CanConstructObject).
    private static bool CanConstructObject(ITypeSymbol type)
        => type.IsValueType ||
           (type is INamedTypeSymbol { TypeKind: TypeKind.Class, IsAbstract: false } named &&
            named.InstanceConstructors.Any(c => c.Parameters.Length == 0 && c.DeclaredAccessibility == Accessibility.Public));

    // The key and value types of a dictionary (see DictionaryRule.TryGetKeyValueTypes).
    private bool TryGetKeyValueTypes(ITypeSymbol type, out ITypeSymbol keyType, out ITypeSymbol valueType)
    {
        if (ImplementsGeneric(type, _enumerableOfT, out var pairType) &&
            pairType is INamedTypeSymbol { IsGenericType: true } pair &&
            SymbolEqualityComparer.Default.Equals(pair.OriginalDefinition, _keyValuePair))
        {
            keyType = pair.TypeArguments[0];
            valueType = pair.TypeArguments[1];
            return true;
        }

        if (Implements(type, _dictionary))
        {
            keyType = _compilation.GetSpecialType(SpecialType.System_String);
            valueType = _compilation.GetSpecialType(SpecialType.System_Object);
            return true;
        }

        if (Implements(type, _enumerable) && GetSingleAdd(type, 2) is { } add)
        {
            keyType = add.Parameters[0].Type;
            valueType = add.Parameters[1].Type;
            return true;
        }

        keyType = valueType = null!;
        return false;
    }

    // The element type of a collection (see ReflectionUtilities.GetElementTypeOrEnumerableType).
    private bool TryGetElementType(ITypeSymbol type, out ITypeSymbol elementType)
    {
        if (ImplementsGeneric(type, _enumerableOfT, out var element))
        {
            elementType = element;
            return true;
        }

        if (Implements(type, _enumerable))
        {
            elementType = type.TypeKind != TypeKind.Interface && GetSingleAdd(type, 1) is { } add && add.Parameters[0].Type.SpecialType != SpecialType.System_Object
                ? add.Parameters[0].Type
                : _compilation.GetSpecialType(SpecialType.System_Object);
            return true;
        }

        elementType = null!;
        return false;
    }

    // The only public instance Add method with the given number of parameters, if there's one.
    private static IMethodSymbol? GetSingleAdd(ITypeSymbol type, int parameterCount)
    {
        IMethodSymbol? result = null;
        foreach (var method in GetAllMembers(type).OfType<IMethodSymbol>())
        {
            if (method is { Name: "Add", IsStatic: false, DeclaredAccessibility: Accessibility.Public } && method.Parameters.Length == parameterCount)
            {
                if (result is not null)
                {
                    return null;
                }

                result = method;
            }
        }

        return result;
    }

    // The public instance properties and fields auto-mapping reads, except those with
    // [ExcelIgnore] (see AutoMapper.TryAddMembers).
    private IEnumerable<ISymbol> GetMappedMembers(ITypeSymbol type)
    {
        var names = new HashSet<string>();
        foreach (var member in GetAllMembers(type))
        {
            if (member.IsStatic || member.DeclaredAccessibility != Accessibility.Public || !names.Add(member.Name) || IsIgnored(member))
            {
                continue;
            }

            switch (member)
            {
                case IPropertySymbol { IsIndexer: false, GetMethod: not null }:
                case IFieldSymbol { IsConst: false, IsImplicitlyDeclared: false }:
                    yield return member;
                    break;
            }
        }
    }

    private bool IsIgnored(ISymbol member)
        => member.GetAttributes().Any(a => SymbolEqualityComparer.Default.Equals(a.AttributeClass, _ignoreAttribute));

    // The type's members and its base types', most derived first.
    private static IEnumerable<ISymbol> GetAllMembers(ITypeSymbol type)
    {
        for (var current = type; current is not null && current.SpecialType != SpecialType.System_Object; current = current.BaseType)
        {
            foreach (var member in current.GetMembers())
            {
                yield return member;
            }
        }
    }

    private static bool Implements(ITypeSymbol type, INamedTypeSymbol? @interface)
        => @interface is not null &&
           (SymbolEqualityComparer.Default.Equals(type, @interface) || type.AllInterfaces.Any(i => SymbolEqualityComparer.Default.Equals(i, @interface)));

    // Whether the type is or implements a closed version of a generic interface, giving its first type argument.
    private static bool ImplementsGeneric(ITypeSymbol type, INamedTypeSymbol? definition, out ITypeSymbol argument)
    {
        if (definition is not null)
        {
            var candidates = type is INamedTypeSymbol { TypeKind: TypeKind.Interface } named ? type.AllInterfaces.Insert(0, named) : type.AllInterfaces;
            foreach (var candidate in candidates)
            {
                if (SymbolEqualityComparer.Default.Equals(candidate.OriginalDefinition, definition))
                {
                    argument = candidate.TypeArguments[0];
                    return true;
                }
            }
        }

        argument = null!;
        return false;
    }

    private static string MetadataName(ITypeSymbol type)
        => type.ContainingNamespace is { IsGlobalNamespace: false } ns ? $"{ns.ToDisplayString()}.{type.MetadataName}" : type.MetadataName;
}
