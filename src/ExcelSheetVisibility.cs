namespace ExcelMapper;

/// <summary>
/// Describes the visibility of an Excel sheet.
/// </summary>
public enum ExcelSheetVisibility
{
    /// <summary>
    /// The sheet is visible.
    /// </summary>
    Visible,

    /// <summary>
    /// The sheet is hidden.
    /// </summary>
    Hidden,

    /// <summary>
    /// The sheet is only visible from VBA code.
    /// </summary>
    VeryHidden
}
