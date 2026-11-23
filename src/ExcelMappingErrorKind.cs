namespace ExcelMapper;

/// <summary>
/// The kind of an <see cref="ExcelMappingError"/>.
/// </summary>
public enum ExcelMappingErrorKind
{
    /// <summary>
    /// A cell is empty, and its member can't be empty.
    /// </summary>
    EmptyValue,

    /// <summary>
    /// A cell's value isn't valid for its member, such as text that isn't a number.
    /// </summary>
    InvalidValue,

    /// <summary>
    /// The row failed data annotation validation (see
    /// <see cref="ExcelImporterConfiguration.ValidateDataAnnotations"/>).
    /// </summary>
    DataAnnotation,

    /// <summary>
    /// Mapping the row threw an <see cref="ExcelMappingException"/>, for example from a
    /// constructor or converter.
    /// </summary>
    Exception,
}
