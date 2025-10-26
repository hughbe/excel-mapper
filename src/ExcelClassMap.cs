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
    public ExcelClassMap(Type type) : this(type, FallbackStrategy.ThrowIfPrimitive)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelClassMap"/> class.
    /// </summary>
    /// <param name="type">The type of the object to create.</param>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    public ExcelClassMap(Type type, FallbackStrategy emptyValueStrategy)
    {
        ArgumentNullException.ThrowIfNull(type);
        EnumUtilities.ValidateIsDefined(emptyValueStrategy);

        Type = type;
        EmptyValueStrategy = emptyValueStrategy;
    }

    /// <summary>
    /// Gets the type of the object to create.
    /// </summary>
    public Type Type { get; }

    internal IMap? _valueMap;

    /// <summary>
    /// Gets the collection of property maps.
    /// </summary>
    public Collection<ExcelPropertyMap> Properties { get; } = new NonNullCollection<ExcelPropertyMap>();

    /// <summary>
    /// Gets the default strategy to use when the value of a cell is empty.
    /// </summary>
    public FallbackStrategy EmptyValueStrategy { get; private protected set; }

    /// <inheritdoc/>
    public virtual bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        // If the map represents a single value, use that map to get the value.
        if (_valueMap is not null)
        {
            return _valueMap.TryGetValue(sheet, rowIndex, reader, member, out result);
        }

        // Otherwise, create an instance of the object and populate its properties.
        var instance = Activator.CreateInstance(Type)!;
        foreach (var property in Properties)
        {
            if (property.Map.TryGetValue(sheet, rowIndex, reader, property.Member, out var value))
            {
                property.SetValueFactory(instance, value);
            }
        }

        result = instance;
        return true;
    }
}
