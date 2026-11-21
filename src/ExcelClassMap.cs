using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// Maps a row of a sheet to an object of the given type.
/// </summary>
public class ExcelClassMap : IMap
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelClassMap"/> class.
    /// </summary>
    /// <param name="type">The type of the object to create.</param>
    public ExcelClassMap([DynamicallyAccessedMembers(TrimmingConstants.MappedType)] Type type) : this(type, FallbackStrategy.ThrowIfPrimitive)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelClassMap"/> class.
    /// </summary>
    /// <param name="type">The type of the object to create.</param>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    public ExcelClassMap([DynamicallyAccessedMembers(TrimmingConstants.MappedType)] Type type, FallbackStrategy emptyValueStrategy)
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

    internal IMap? _valueMap;

    /// <summary>
    /// Gets the collection of property maps.
    /// </summary>
    public Collection<ExcelPropertyMap> Properties { get; } = new PropertyMapCollection();

    /// <summary>
    /// Gets the default strategy to use when the value of a cell is empty.
    /// </summary>
    public FallbackStrategy EmptyValueStrategy { get; private protected set; }

    private ConstructionPlan? _constructionPlan;

    /// <inheritdoc/>
    public virtual bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        // If the map represents a single value, use that map to get the value.
        if (_valueMap is not null)
        {
            if (sheet.CollectedErrors is not { } valueErrors)
            {
                return _valueMap.TryGetValue(sheet, rowIndex, reader, member, out result);
            }

            try
            {
                return _valueMap.TryGetValue(sheet, rowIndex, reader, member, out result);
            }
            catch (ExcelMappingException exception)
            {
                valueErrors.Add(exception);
                result = null;
                return false;
            }
        }

        var plan = GetConstructionPlan();
        if (plan.Constructor is null)
        {
            // Create an instance of the object and populate its properties.
            var instance = plan.CreateInstance!();
            var properties = Properties;
            for (int i = 0; i < properties.Count; i++)
            {
                var property = properties[i];
                if (TryGetMemberValue(property, sheet, rowIndex, reader, out var value))
                {
                    property.SetValueFactory(instance, value!);
                }
            }

            result = instance;
            return true;
        }

        return TryCreateUsingConstructor(plan, sheet, rowIndex, reader, out result);
    }

    /// <summary>
    /// Gets the value of a member. When the sheet is collecting errors (see
    /// <see cref="ExcelSheet.ReadRowResults{T}"/>), a mapping error is recorded and the member is
    /// skipped so that the remaining members are still read.
    /// </summary>
    private static bool TryGetMemberValue(ExcelPropertyMap property, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out object? value)
    {
        if (sheet.CollectedErrors is not { } errors)
        {
            return property.Map.TryGetValue(sheet, rowIndex, reader, property.Member, out value);
        }

        var errorCount = errors.Count;
        try
        {
            if (property.Map.TryGetValue(sheet, rowIndex, reader, property.Member, out value) && errors.Count == errorCount)
            {
                return true;
            }
        }
        catch (ExcelMappingException exception)
        {
            errors.Add(exception);
        }

        // An error was thrown, or recorded without throwing (see ThrowFallback): don't set the member.
        value = null;
        return false;
    }

    private bool TryCreateUsingConstructor(ConstructionPlan plan, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, [NotNullWhen(true)] out object? result)
    {
        var errors = sheet.CollectedErrors;
        var errorCount = errors?.Count ?? 0;
        // Read every member first, as constructor arguments must be known before the object exists.
        var properties = Properties;
        var count = properties.Count;
        var arguments = (object?[])plan.DefaultArguments.Clone();

        // Values for members set after construction. Only allocated if there are any, which there
        // usually aren't for records.
        object?[]? values = null;
        bool[]? hasValues = null;
        for (int i = 0; i < count; i++)
        {
            var property = properties[i];
            if (TryGetMemberValue(property, sheet, rowIndex, reader, out var value))
            {
                var parameterIndex = plan.PropertyParameterIndices[i];
                if (parameterIndex >= 0)
                {
                    arguments[parameterIndex] = value;
                }
                else
                {
                    values ??= new object?[count];
                    hasValues ??= new bool[count];
                    values[i] = value;
                    hasValues[i] = true;
                }
            }
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
            instance = plan.InvokeConstructor!(arguments);
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
        if (hasValues is not null)
        {
            for (int i = 0; i < count; i++)
            {
                if (hasValues[i])
                {
                    properties[i].SetValueFactory(instance, values![i]!);
                }
            }
        }

        result = instance;
        return true;
    }

    private ConstructionPlan GetConstructionPlan()
    {
        var properties = (PropertyMapCollection)Properties;
        var plan = _constructionPlan;
        if (plan is null || plan.Version != properties.Version)
        {
            plan = CreateConstructionPlan(properties.Version);
            _constructionPlan = plan;
        }

        return plan;
    }

    private ConstructionPlan CreateConstructionPlan(int version)
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
            return new ConstructionPlan(version, null, [], []) { CreateInstance = CompiledAccessors.GetParameterlessConstructor(Type) };
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

        return new ConstructionPlan(version, constructor, propertyParameterIndices, defaultArguments) { InvokeConstructor = CompiledAccessors.GetConstructor(constructor) };
    }

    private sealed class ConstructionPlan(int version, ConstructorInfo? constructor, int[] propertyParameterIndices, object?[] defaultArguments)
    {
        public int Version { get; } = version;

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
        /// Calls <see cref="Constructor"/> with an array of arguments.
        /// </summary>
        public Func<object?[], object>? InvokeConstructor { get; set; }
    }

    /// <summary>
    /// A collection of property maps that tracks changes so the construction plan can be updated.
    /// </summary>
    private sealed class PropertyMapCollection : NonNullCollection<ExcelPropertyMap>
    {
        public int Version { get; private set; }

        protected override void InsertItem(int index, ExcelPropertyMap item)
        {
            base.InsertItem(index, item);
            Version++;
        }

        protected override void SetItem(int index, ExcelPropertyMap item)
        {
            base.SetItem(index, item);
            Version++;
        }

        protected override void RemoveItem(int index)
        {
            base.RemoveItem(index);
            Version++;
        }

        protected override void ClearItems()
        {
            base.ClearItems();
            Version++;
        }
    }
}
