namespace ExcelMapper.Abstractions;

/// <summary>
/// Cell mapper that supports formats.
/// </summary>
internal interface IFormatsCellMapper : IFormatProviderCellMapper
{
    /// <summary>
    /// The formats supported by the cell mapper.
    /// </summary>
    string[] Formats { get; set; }
}
