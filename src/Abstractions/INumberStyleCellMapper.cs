using System.Globalization;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Cell mapper that supports number styles.
/// </summary>
internal interface INumberStyleCellMapper : IFormatProviderCellMapper
{
    /// <summary>
    /// The number styles used for parsing.
    /// </summary>
    NumberStyles Style { get; set; }
}
