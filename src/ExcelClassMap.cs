using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// Maps a row of a sheet to an object of the given type. Class maps derive from
/// <see cref="ExcelClassMap{T}"/>, which has the methods that map members: this is its base, so
/// that class maps of any type can be registered and looked up. Only <see cref="ExcelClassMap{T}"/>
/// can derive from it, as it has abstract members that are internal to the library.
/// </summary>
public abstract class ExcelClassMap : IParentMap, IBindableMap
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelClassMap"/> class.
    /// </summary>
    /// <param name="type">The type of the object to create.</param>
    protected ExcelClassMap([DynamicallyAccessedMembers(TrimmingConstants.MappedType)] Type type) : this(type, EmptyValueStrategy.Error)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelClassMap"/> class.
    /// </summary>
    /// <param name="type">The type of the object to create.</param>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    protected ExcelClassMap([DynamicallyAccessedMembers(TrimmingConstants.MappedType)] Type type, EmptyValueStrategy emptyValueStrategy)
    {
        ThrowHelpers.ThrowIfNull(type, nameof(type));
        EnumUtilities.ValidateIsDefined(emptyValueStrategy);

        Type = type;
        EmptyValueStrategy = emptyValueStrategy;
    }

    /// <summary>
    /// Gets the type of the object to create.
    /// </summary>
    [DynamicallyAccessedMembers(TrimmingConstants.MappedType)]
    public Type Type { get; }

    /// <summary>
    /// Gets or sets the map of a class map read as a single value (<c>Map(p => p)</c>).
    /// </summary>
    internal IMap? SingleValueMap
    {
        get => _valueMap;
        set
        {
            ThrowHelpers.ThrowIfFrozen(_frozen);
            _valueMap = value;
        }
    }

    private IMap? _valueMap;

    private bool _frozen;

    /// <inheritdoc />
    void IMap.Freeze() => Freeze();

    /// <summary>
    /// Stops the class map and its members' maps changing, the first time rows are read with it.
    /// </summary>
    internal void Freeze()
    {
        if (_frozen)
        {
            return;
        }

        // A class map that maps nothing would give an empty object for every row, which is
        // never what's meant: usually the type's members are read-only or not public.
        if (_valueMap is null && Members.Count == 0)
        {
            throw new ExcelMappingException($"The class map for type {TypeNames.Get(Type)} maps no members, so every row would be an empty object. Auto-mapping maps public fields, and public properties with a public setter or set by a public constructor; map other members with Map(...) in a class map.");
        }

        _frozen = true;
        Members.Freeze();
        _valueMap?.Freeze();
        foreach (var member in Members)
        {
            member.Map.Freeze();
        }
    }

    IEnumerable<KeyValuePair<object?, IMap>> IParentMap.Children
    {
        get
        {
            // A class map read as a single value (Map(p => p)) has a child with no key.
            if (_valueMap is not null)
            {
                yield return new(null, _valueMap);
            }

            foreach (var member in Members)
            {
                yield return new(member.Member, member.Map);
            }
        }
    }

    IMap? IParentMap.GetChild(object key)
    {
        foreach (var member in Members)
        {
            if (member.Member.Equals(key))
            {
                return member.Map;
            }
        }

        return null;
    }

    void IParentMap.SetChild(object key, IMap map)
    {
        var member = (MemberInfo)key;
        for (int i = 0; i < Members.Count; i++)
        {
            if (Members[i].Member.Equals(member))
            {
                Members.RemoveAt(i);
                break;
            }
        }

        Members.Add(new ClassMapMember(member, map));
    }

    string? IParentMap.GetChildPath(string? path, object? key)
        => key is MemberInfo member ? (path is null ? member.Name : $"{path}.{member.Name}") : path;

    /// <summary>
    /// Gets the members to map, each with the map that reads it. They can't change once the class
    /// map is frozen.
    /// </summary>
    internal NonNullCollection<ClassMapMember> Members { get; } = [];

    /// <summary>
    /// Gets the default strategy to use when the value of a cell is empty.
    /// </summary>
    public EmptyValueStrategy EmptyValueStrategy { get; private protected set; }

    /// <inheritdoc/>
    bool IMap.TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? result)
        => TryGetObject(sheet, rowIndex, reader, member, errors, out result);

    /// <summary>
    /// Maps a row (or for a nested object, the cells it reads) to an object (see
    /// <see cref="ExcelClassMap{T}.TryGetValue"/>).
    /// </summary>
    private protected abstract bool TryGetObject(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? result);

    /// <inheritdoc/>
    BoundMember<TOwner> IBindableMap.Bind<TOwner>(ClassMapMember member) => Bind<TOwner>(member);

    /// <inheritdoc/>
    MethodInfo? IBindableMap.GetArgumentMethod<TOwner>(Type parameterType) => GetArgumentMethod<TOwner>(parameterType);

    // Binds the class map to a member of another class map, as a nested object.
    private protected abstract BoundMember<TOwner> Bind<TOwner>(ClassMapMember member);

    private protected abstract MethodInfo? GetArgumentMethod<TOwner>(Type parameterType);
}
