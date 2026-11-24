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
/// <c>ReflectionUtilities.GetElementTypeOrEnumerableType</c>, <c>GetMappedMembers</c> follows
/// <c>AutoMapper.TryAddMembers</c>, <c>AddSetter</c> follows <c>CompiledAccessors.GetSetter</c>, and
/// <c>AddConstructors</c> follows <c>ConstructionPlan&lt;T&gt;.Create</c>. See "Auto-mapping and the source generator" in CLAUDE.md.</para>
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

    // Attributes that give a member's elements a conversion, or make it read something other than
    // cells (see AutoMapper.TryAutoMapAndAddMember and MemberMapper.AddMappers).
    private readonly INamedTypeSymbol?[] _elementAttributes;
    private readonly bool _canRegisterParsable;

    private readonly HashSet<ITypeSymbol> _visited = new(SymbolEqualityComparer.Default);
    private readonly List<string> _registrations = [];
    private readonly HashSet<string> _registered = [];

    // The object types whose members are being visited, to find a type that contains itself.
    private readonly HashSet<ITypeSymbol> _objectPath = new(SymbolEqualityComparer.Default);

    // What will make auto-mapping fail for the row being added (see TakeProblems).
    private readonly List<(DiagnosticDescriptor Descriptor, string[] Arguments)> _problems = [];
    private string _rowName = "";

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
        _elementAttributes =
        [
            compilation.GetTypeByMetadataName("ExcelMapper.ExcelConverterAttribute"),
            compilation.GetTypeByMetadataName("ExcelMapper.ExcelMappingDictionaryAttribute"),
            compilation.GetTypeByMetadataName("ExcelMapper.ExcelRowNumberAttribute"),
            compilation.GetTypeByMetadataName("ExcelMapper.ExcelExtraColumnsAttribute"),
        ];

        // The library's netstandard builds don't have IParsable<T> registrations.
        _canRegisterParsable = compilation.GetTypeByMetadataName("ExcelMapper.ExcelTypeRegistrar")?.GetMembers("Parsable").Length > 0;
    }

    public IReadOnlyList<string> Registrations => _registrations;

    /// <summary>
    /// Gets what will make auto-mapping a row type fail, found while adding it: a type with nothing
    /// to set, a type that can't be created, or a type that contains itself. Each would make
    /// <c>GetClassMap</c> throw for the row type.
    /// </summary>
    public List<(DiagnosticDescriptor Descriptor, string[] Arguments)> TakeProblems()
    {
        var problems = new List<(DiagnosticDescriptor, string[])>(_problems);
        _problems.Clear();
        return problems;
    }

    /// <summary>
    /// Registers a type rows are mapped to, and every type it uses.
    /// </summary>
    public void AddRow(ITypeSymbol type)
    {
        _rowName = Display(type);
        Add($"Row<{Name(type)}>()");

        // Row registers the type as Type does.
        _registered.Add($"Type<{Name(type)}>()");
        CheckElements(type, _rowName);
        Visit(type);
    }

    // A type's name in a message, as written in C#, without a nullable annotation.
    private static string Display(ITypeSymbol type) => type.WithNullableAnnotation(NullableAnnotation.NotAnnotated).ToDisplayString();

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
        if (type.TypeKind == TypeKind.TypeParameter)
        {
            return;
        }

        CheckObject(type);
        _objectPath.Add(type);
        foreach (var member in GetMappedMembers(type))
        {
            var memberType = MemberType(member);
            AddSetter(member, memberType, name);

            // A member of the type being mapped, or of a type it's in, makes auto-mapping recurse
            // forever (see AutoMapper.TryCreateObjectMap, which throws).
            var objectType = memberType is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullableMember ? nullableMember.TypeArguments[0] : memberType;
            if (_objectPath.Contains(objectType) && IsAutoMapped(type, member))
            {
                _problems.Add((Descriptors.ContainsItself, [Display(objectType), $"{Display(type)}.{member.Name}", _rowName]));
                continue;
            }

            if (IsAutoMapped(type, member) && !member.GetAttributes().Any(a => _elementAttributes.Contains(a.AttributeClass, SymbolEqualityComparer.Default)))
            {
                CheckElements(memberType, $"{Display(type)}.{member.Name}");
            }

            Visit(memberType);
        }

        _objectPath.Remove(type);
    }

    // Reports why auto-mapping can't map an object type, if it can't: it can't create it (see
    // AutoMapper.CanConstructObjectMap), or it has nothing to set (see ExcelClassMap.Freeze).
    private void CheckObject(ITypeSymbol type)
    {
        if (type.TypeKind == TypeKind.Interface || (type.IsAbstract && !type.IsValueType))
        {
            _problems.Add((Descriptors.CannotCreate, [Display(type), _rowName, "it's an interface or abstract class: map a member of a type that can be created"]));
            return;
        }

        if (!CanConstructObject(type) && (Implements(type, _enumerable) || !HasConstructorFromMembers(type)))
        {
            _problems.Add((Descriptors.CannotCreate, [Display(type), _rowName, "it needs a public parameterless constructor, or a public constructor whose parameters each match a property or field (by name and type) or have a default value"]));
            return;
        }

        if (!GetMappedMembers(type).Any(m => IsAutoMapped(type, m)))
        {
            _problems.Add((Descriptors.NothingToMap, [Display(type), _rowName]));
        }
    }

    // Reports a collection or dictionary whose elements can't be read from a cell, such as
    // List<Address>: auto-mapping throws rather than map it as an object (see AutoMapper.TryCreateMap).
    private void CheckElements(ITypeSymbol type, string where)
    {
        if (IsReadable(type))
        {
            return;
        }

        var isDictionary = TryGetKeyValueTypes(type, out _, out var valueType);
        var isCollection = TryGetElementType(type, out var elementType);
        if ((isDictionary || isCollection) && !(isDictionary && IsReadable(valueType)) && !(isCollection && IsReadable(elementType)))
        {
            _problems.Add((Descriptors.ElementsNotReadable, [where, Display(isDictionary ? valueType : elementType), _rowName]));
        }
    }

    // Whether a value of the type is read from a cell: it has a built-in conversion (see
    // TypeRule<T>.FindDefaultMappers). A type parameter isn't known, so it's assumed to be.
    private bool IsReadable(ITypeSymbol type) => type switch
    {
        INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable => IsReadable(nullable.TypeArguments[0]),
        { TypeKind: TypeKind.Enum or TypeKind.TypeParameter or TypeKind.Error } => true,
        _ => s_builtInTypes.Contains(MetadataName(type)) || (HasConversion(type, out var isParsable) && (!isParsable || _canRegisterParsable)),
    };

    // Whether auto-mapping maps a member GetMappedMembers finds (see AutoMapper.TryAddMembers and
    // ShouldAutoMapProperty): a field, a property with a public setter, or a read-only property
    // set through a constructor parameter.
    private bool IsAutoMapped(ITypeSymbol type, ISymbol member) => member switch
    {
        IFieldSymbol => true,
        IPropertySymbol { SetMethod: { } setter } => setter.DeclaredAccessibility == Accessibility.Public,
        IPropertySymbol property => property.GetMethod is { DeclaredAccessibility: Accessibility.Public } &&
            !Implements(type, _enumerable) &&
            PublicConstructors(type).Any(c => c.Parameters.Any(p => ParameterMatches(p, property.Name, property.Type))),
        _ => false,
    };

    // Whether a public constructor's parameters each match a property or field the type has, or
    // have a default value (see ConstructorUtilities.FindConstructor).
    private bool HasConstructorFromMembers(ITypeSymbol type)
    {
        var readable = new List<(string Name, ITypeSymbol Type)>();
        foreach (var member in GetAllMembers(type))
        {
            switch (member)
            {
                case IPropertySymbol { IsStatic: false, IsIndexer: false, GetMethod.DeclaredAccessibility: Accessibility.Public } property:
                    readable.Add((property.Name, property.Type));
                    break;
                case IFieldSymbol { IsStatic: false, IsConst: false, DeclaredAccessibility: Accessibility.Public } field:
                    readable.Add((field.Name, field.Type));
                    break;
            }
        }

        return PublicConstructors(type).Any(c => c.Parameters.Length > 0 &&
            c.Parameters.All(p => p.HasExplicitDefaultValue || readable.Any(m => ParameterMatches(p, m.Name, m.Type))));
    }

    private static IEnumerable<IMethodSymbol> PublicConstructors(ITypeSymbol type)
        => type is INamedTypeSymbol named ? named.InstanceConstructors.Where(c => c.DeclaredAccessibility == Accessibility.Public) : [];

    // A parameter with the member's name, ignoring case, that the member's value can be passed to
    // as it is (see ConstructorUtilities.ParameterMatches, which uses Type.IsAssignableFrom).
    private bool ParameterMatches(IParameterSymbol parameter, string name, ITypeSymbol memberType)
    {
        if (!string.Equals(parameter.Name, name, System.StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var conversion = _compilation.ClassifyCommonConversion(memberType, parameter.Type);
        return conversion.Exists && conversion.IsImplicit && !conversion.IsNumeric && !conversion.IsUserDefined;
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
