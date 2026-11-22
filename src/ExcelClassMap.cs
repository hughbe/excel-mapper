using System.Collections.ObjectModel;
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
    public ExcelClassMap([DynamicallyAccessedMembers(TrimmingConstants.MappedType)] Type type) : this(type, EmptyValueStrategy.Error)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelClassMap"/> class.
    /// </summary>
    /// <param name="type">The type of the object to create.</param>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    public ExcelClassMap([DynamicallyAccessedMembers(TrimmingConstants.MappedType)] Type type, EmptyValueStrategy emptyValueStrategy)
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
        ((PropertyMapCollection)Properties).Freeze();
        _valueMap?.Freeze();
        foreach (var property in Properties)
        {
            property.Map.Freeze();
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

            foreach (var property in Properties)
            {
                yield return new(property.Member, property.Map);
            }
        }
    }

    IMap? IParentMap.GetChild(object key)
    {
        foreach (var property in Properties)
        {
            if (property.Member.Equals(key))
            {
                return property.Map;
            }
        }

        return null;
    }

    void IParentMap.SetChild(object key, IMap map)
    {
        var member = (MemberInfo)key;
        for (int i = 0; i < Properties.Count; i++)
        {
            if (Properties[i].Member.Equals(member))
            {
                Properties.RemoveAt(i);
                break;
            }
        }

        Properties.Add(new ExcelPropertyMap(member, map));
    }

    string? IParentMap.GetChildPath(string? path, object? key)
        => key is MemberInfo member ? (path is null ? member.Name : $"{path}.{member.Name}") : path;

    /// <summary>
    /// Gets the collection of property maps.
    /// </summary>
    internal Collection<ExcelPropertyMap> Properties { get; } = new PropertyMapCollection();

    /// <summary>
    /// Gets the default strategy to use when the value of a cell is empty.
    /// </summary>
    public EmptyValueStrategy EmptyValueStrategy { get; private protected set; }

    private ConstructionPlan? _constructionPlan;

    /// <inheritdoc/>
    bool IMap.TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? result)
        => TryGetValue(sheet, rowIndex, reader, member, errors, out result);

    /// <summary>
    /// Maps a row (or for a nested object, the cells it reads) to an object. When
    /// <paramref name="errors"/> isn't null (see <see cref="ExcelSheet.ReadRowResults{T}"/>), mapping
    /// errors are added to it rather than thrown, and every member is still read.
    /// </summary>
    internal bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? result)
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

    private static void MapBoundMembers(BoundMember[] members, object instance, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
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
    private static void ReadMemberValue(BoundMember member, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors)
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

    private bool TryCreateUsingConstructor(ConstructionPlan plan, BoundMember[] members, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? result)
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
            instance = plan.CreateFromMembers is { } createFromMembers
                ? createFromMembers(members)
                : plan.InvokeConstructor!(GetArguments(plan, members));
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

        // Set any members that weren't passed to the constructor.
        var properties = Properties;
        for (int i = 0; i < members.Length; i++)
        {
            if (plan.PropertyParameterIndices[i] < 0 && members[i].HasValue)
            {
                members[i].SetValue(instance, properties[i]);
            }
        }

        result = instance;
        return true;
    }

    // Where constructors can't be compiled: the constructor's arguments as an array.
    private static object?[] GetArguments(ConstructionPlan plan, BoundMember[] members)
    {
        var arguments = (object?[])plan.DefaultArguments.Clone();
        for (int i = 0; i < members.Length; i++)
        {
            var parameterIndex = plan.PropertyParameterIndices[i];
            if (parameterIndex >= 0 && members[i].HasValue)
            {
                arguments[parameterIndex] = members[i].Value;
            }
        }

        return arguments;
    }

    // Made when the first row is read, after the class map is frozen.
    private ConstructionPlan GetConstructionPlan() => _constructionPlan ??= CreateConstructionPlan();

    private ConstructionPlan CreateConstructionPlan()
    {
        var hasReadOnlyMember = false;
        foreach (var property in Properties)
        {
            if (property.IsReadOnly)
            {
                hasReadOnlyMember = true;
                break;
            }
        }

        // Prefer a parameterless constructor and setting members, unless a member can only be
        // set through a constructor.
        if (!hasReadOnlyMember && ConstructorUtilities.HasParameterlessConstructor(Type))
        {
            return new ConstructionPlan(null, [], []) { CreateInstance = CompiledAccessors.GetParameterlessConstructor(Type) };
        }

        if (Type.IsInterface || Type.IsAbstract)
        {
            throw new ExcelMappingException($"Cannot create an instance of type \"{Type}\" as it is an interface or abstract class.");
        }

        var members = new MemberInfo[Properties.Count];
        for (int i = 0; i < members.Length; i++)
        {
            members[i] = Properties[i].Member;
        }

        var constructor = ConstructorUtilities.FindConstructor(Type, members, out var parameterMemberIndices);
        if (constructor is null)
        {
            throw new ExcelMappingException(
                $"Cannot create an instance of type \"{Type}\". It needs a public parameterless constructor, or a public " +
                "constructor whose parameters each match a mapped property or field (by name and type) or have a default value.");
        }

        var propertyParameterIndices = new int[members.Length];
        for (int i = 0; i < propertyParameterIndices.Length; i++)
        {
            propertyParameterIndices[i] = -1;
        }
        var parameters = constructor.GetParameters();
        var defaultArguments = new object?[parameters.Length];
        for (int i = 0; i < parameters.Length; i++)
        {
            if (parameterMemberIndices[i] >= 0)
            {
                propertyParameterIndices[parameterMemberIndices[i]] = i;
            }

            // Reflection passes a zero-initialized value for null value type arguments.
            defaultArguments[i] = parameters[i].HasDefaultValue ? parameters[i].DefaultValue : null;
        }

        for (int i = 0; i < members.Length; i++)
        {
            if (propertyParameterIndices[i] == -1 && Properties[i].IsReadOnly)
            {
                throw new ExcelMappingException($"Property \"{members[i].Name}\" of type \"{Type}\" is read-only and is not a parameter of the constructor {constructor}.");
            }
        }

        var boundArguments = new BoundConstructorArgument[parameters.Length];
        for (int i = 0; i < parameters.Length; i++)
        {
            var memberIndex = parameterMemberIndices[i];
            var typedMethod = memberIndex >= 0 && Properties[memberIndex].Map is IValueMap bindableMap
                ? bindableMap.GetArgumentMethod(parameters[i].ParameterType)
                : null;
            boundArguments[i] = new BoundConstructorArgument(memberIndex, typedMethod, defaultArguments[i]);
        }

        var createFromMembers = CompiledAccessors.GetConstructor(constructor, boundArguments);
        return new ConstructionPlan(constructor, propertyParameterIndices, defaultArguments)
        {
            CreateFromMembers = createFromMembers,
            InvokeConstructor = createFromMembers is null ? CompiledAccessors.GetConstructor(constructor) : null,
        };
    }

    private sealed class ConstructionPlan(ConstructorInfo? constructor, int[] propertyParameterIndices, object?[] defaultArguments)
    {
        /// <summary>
        /// The constructor to call, or null to use the parameterless constructor.
        /// </summary>
        public ConstructorInfo? Constructor { get; } = constructor;

        /// <summary>
        /// For each property map, the index of the constructor parameter it is passed to, or -1.
        /// </summary>
        public int[] PropertyParameterIndices { get; } = propertyParameterIndices;

        public object?[] DefaultArguments { get; } = defaultArguments;

        /// <summary>
        /// Creates an instance using the parameterless constructor, if <see cref="Constructor"/> is null.
        /// </summary>
        public Func<object>? CreateInstance { get; set; }

        /// <summary>
        /// Calls <see cref="Constructor"/> with the values the bound members read, or null where
        /// code can't be compiled.
        /// </summary>
        public Func<BoundMember[], object>? CreateFromMembers { get; set; }

        /// <summary>
        /// Calls <see cref="Constructor"/> with an array of arguments, where code can't be compiled.
        /// </summary>
        public Func<object?[], object>? InvokeConstructor { get; set; }
    }

    /// <summary>
    /// A collection of property maps, which can't change once the class map is frozen.
    /// </summary>
    private sealed class PropertyMapCollection : NonNullCollection<ExcelPropertyMap>
    {
    }
}
