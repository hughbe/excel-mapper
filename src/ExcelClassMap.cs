using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// Maps a row of a sheet to an object of the given type.
/// </summary>
public class ExcelClassMap : IParentMap
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelClassMap"/> class.
    /// </summary>
    /// <param name="type">The type of the object to create.</param>
    private protected ExcelClassMap([DynamicallyAccessedMembers(TrimmingConstants.MappedType)] Type type) : this(type, EmptyValueStrategy.Error)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelClassMap"/> class.
    /// </summary>
    /// <param name="type">The type of the object to create.</param>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    private protected ExcelClassMap([DynamicallyAccessedMembers(TrimmingConstants.MappedType)] Type type, EmptyValueStrategy emptyValueStrategy)
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

    private ConstructionPlan? _constructionPlan;

    /// <inheritdoc/>
    bool IMap.TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? result)
        => TryGetValue(sheet, rowIndex, reader, member, errors, out result);

    /// <summary>
    /// Maps a row (or for a nested object, the cells it reads) to an object. When
    /// <paramref name="errors"/> isn't null (see <see cref="ExcelSheet.ReadRowResults{T}"/>), mapping
    /// errors are added to it rather than thrown, and every member is still read.
    /// </summary>
    internal bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? result)
    {
        // If the map represents a single value, use that map to get the value.
        if (_valueMap is not null)
        {
            if (errors is null)
            {
                return _valueMap.TryGetValue(sheet, rowIndex, reader, member, null, out result);
            }

            try
            {
                return _valueMap.TryGetValue(sheet, rowIndex, reader, member, errors, out result);
            }
            catch (ExcelMappingException exception)
            {
                errors.Add(exception);
                result = null;
                return false;
            }
        }

        var plan = GetConstructionPlan();
        var members = sheet.GetBoundMembers(this);
        if (plan.Constructor is null)
        {
            // Create an instance of the object and populate its properties.
            var instance = plan.CreateInstance!();
            MapBoundMembers(members, instance, sheet, rowIndex, reader, errors);
            result = instance;
            return true;
        }

        return TryCreateUsingConstructor(plan, members, sheet, rowIndex, reader, errors, out result);
    }

    private static void MapBoundMembers(BoundMember[] members, object instance, SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        if (errors is null)
        {
            for (int i = 0; i < members.Length; i++)
            {
                members[i].TryMap(instance, sheet, rowIndex, reader, null);
            }

            return;
        }

        for (int i = 0; i < members.Length; i++)
        {
            try
            {
                members[i].TryMap(instance, sheet, rowIndex, reader, errors);
            }
            catch (ExcelMappingException exception)
            {
                errors.Add(exception);
            }
        }
    }

    /// <summary>
    /// Reads a member's value for the row, for a type created through a constructor (see
    /// <see cref="BoundMember.ReadValue"/>). When errors are being collected (see
    /// <see cref="ExcelSheet.ReadRowResults{T}"/>), a mapping error is recorded and the member has no
    /// value, so that the remaining members are still read.
    /// </summary>
    private static void ReadMemberValue(BoundMember member, SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
    {
        if (errors is null)
        {
            member.ReadValue(sheet, rowIndex, reader, null);
            return;
        }

        try
        {
            member.ReadValue(sheet, rowIndex, reader, errors);
        }
        catch (ExcelMappingException exception)
        {
            errors.Add(exception);
        }
    }

    private bool TryCreateUsingConstructor(ConstructionPlan plan, BoundMember[] members, SheetView sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? result)
    {
        var errorCount = errors?.Count ?? 0;

        // Read every member first, as constructor arguments must be known before the object exists.
        // Each member keeps its value until it's passed to the constructor or set.
        for (int i = 0; i < members.Length; i++)
        {
            ReadMemberValue(members[i], sheet, rowIndex, reader, errors);
        }

        // When collecting errors, don't call the constructor with arguments that couldn't be read.
        if (errors is not null && errors.Count > errorCount)
        {
            result = null;
            return false;
        }

        object instance;
        try
        {
            instance = plan.Construct(members);
        }
        catch (Exception ex)
        {
            var inner = ex is TargetInvocationException { InnerException: { } invocationInner } ? invocationInner : ex;
            var exception = new ExcelMappingException($"Could not create an instance of type \"{Type}\": {inner.Message}", sheet, rowIndex, -1, inner);
            if (errors is null)
            {
                throw exception;
            }

            errors.Add(exception);
            result = null;
            return false;
        }

        plan.SetOtherMembers(instance, members);
        result = instance;
        return true;
    }

    // Made when the first row is read, after the class map is frozen.
    private ConstructionPlan GetConstructionPlan() => _constructionPlan ??= ConstructionPlan.Create(Type, Members);
}
