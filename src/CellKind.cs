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
}
