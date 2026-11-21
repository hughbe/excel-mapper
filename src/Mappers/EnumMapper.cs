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
        ThrowHelpers.ThrowIfNull(enumType, nameof(enumType));

        if (!enumType.GetTypeInfo().IsEnum)
        {
            throw new ArgumentException($"Type {enumType} is not an Enum.", nameof(enumType));
        }

        EnumType = enumType;
        IgnoreCase = ignoreCase;
    }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        var enumType = EnumType;
        var ignoreCase = IgnoreCase;
#if NETSTANDARD2_0
        try
        {
            return CellMapperResult.Success(Enum.Parse(enumType, stringValue!, ignoreCase));
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
#else
        if (Enum.TryParse(enumType, stringValue, ignoreCase, out var result))
        {
            return CellMapperResult.Success(result!);
        }

        // Only create the exception (by parsing again) if it is needed.
        return InvalidResult(enumType, stringValue, ignoreCase);
#endif
    }

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(Type enumType, string? stringValue, bool ignoreCase)
        => MapperUtilities.InvalidFromParse(() => Enum.Parse(enumType, stringValue!, ignoreCase));
}
