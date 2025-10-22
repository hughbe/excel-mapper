using System.Reflection;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to an enum of a given type.
/// </summary>
public class EnumMapper : ICellMapper
{
    /// <summary>
    /// Gets the type of the enum to map the value of a cell to.
    /// </summary>
    public Type EnumType { get; }

    /// <summary>
    /// Gets whether enum parsing is case insensitive.
    /// </summary>
    public bool IgnoreCase { get; }

    /// <summary>
    /// Constructs a mapper that tries to map the value of a cell to an enum of a given type.
    /// </summary>
    /// <param name="enumType">The type of the enum to convert the value of a cell to.</param>
    public EnumMapper(Type enumType) : this(enumType, ignoreCase: false)
    {
    }

    /// <summary>
    /// Constructs a mapper that tries to map the value of a cell to an enum of a given type.
    /// </summary>
    /// <param name="enumType">The type of the enum to convert the value of a cell to.</param>
    /// <param name="ignoreCase">A flag indicating whether enum parsing is case insensitive.</param>
    public EnumMapper(Type enumType, bool ignoreCase)
    {
        ArgumentNullException.ThrowIfNull(enumType);

        if (!enumType.GetTypeInfo().IsEnum)
        {
            throw new ArgumentException($"Type {enumType} is not an Enum.", nameof(enumType));
        }

        EnumType = enumType;
        IgnoreCase = ignoreCase;
    }

    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        try
        {
            var result = Enum.Parse(EnumType, stringValue!, IgnoreCase);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}
