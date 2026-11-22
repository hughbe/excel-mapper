namespace ExcelMapper;

/// <summary>
/// Maps rows to a type with an <see cref="ExcelMappingContext"/>, whose registrations the
/// ExcelMapper source generator writes: put it on a <c>partial</c> class deriving from
/// <see cref="ExcelMappingContext"/>, once for each type rows are mapped to. The generator
/// registers the type and every type its members use, adds a property for its class map, and a
/// static <c>Default</c> instance of the context.
/// </summary>
/// <example>
/// <code>
/// [ExcelMappable(typeof(Employee))]
/// public partial class AppMappingContext : ExcelMappingContext
/// {
/// }
///
/// var employees = sheet.ReadRows(AppMappingContext.Default.Employee);
/// </code>
/// </example>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
public sealed class ExcelMappableAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelMappableAttribute"/>.
    /// </summary>
    /// <param name="type">The type rows are mapped to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="type"/> is null.</exception>
    public ExcelMappableAttribute(Type type)
    {
        ThrowHelpers.ThrowIfNull(type, nameof(type));
        Type = type;
    }

    /// <summary>
    /// Gets the type rows are mapped to.
    /// </summary>
    public Type Type { get; }
}
