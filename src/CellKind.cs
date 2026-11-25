namespace ExcelMapper;

/// <summary>
/// The kind of value a cell holds. See <see cref="CellValue.Kind"/>.
/// </summary>
public enum CellKind
{
    /// <summary>The cell is empty, or holds empty text.</summary>
    Empty,

    /// <summary>The cell holds text. CSV cells are always text.</summary>
    Text,

    /// <summary>The cell holds a number. Excel stores every number as a <see cref="double"/>.</summary>
    Number,

    /// <summary>The cell holds <see langword="true"/> or <see langword="false"/>.</summary>
    Boolean,

    /// <summary>The cell holds a date, or a date and time.</summary>
    DateTime,

    /// <summary>The cell holds a duration.</summary>
    TimeSpan,

    /// <summary>The cell holds another kind of value: see <see cref="CellValue.Value"/>.</summary>
    Other,

    /// <summary>
    /// The cell holds a formula error, such as <c>#DIV/0!</c> or <c>#N/A</c>: its
    /// <see cref="CellValue.Text"/> and <see cref="CellValue.Value"/> are the error's text. Errors
    /// are invalid for every type, so they're given to a member's invalid fallback rather than
    /// converted.
    /// </summary>
    Error,
}
