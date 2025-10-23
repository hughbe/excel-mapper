namespace ExcelMapper.Abstractions;

/// <summary>
/// Cell mapper that supports formats.
/// </summary>
internal interface IFormatsCellMapper : ICellMapper
{
    /// <summary>
    /// The formats supported by the cell mapper.
    /// </summary>
    string[] Formats { get; set; }

    /// <summary>
    /// The format provider used for formatting.
    /// </summary>
    IFormatProvider? Provider { get; set; }
}
