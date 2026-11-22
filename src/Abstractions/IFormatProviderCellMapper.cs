namespace ExcelMapper.Abstractions;

/// <summary>
/// Cell mapper that supports format provider.
/// </summary>
internal interface IFormatProviderCellMapper
{
    /// <summary>
    /// The format provider used for formatting.
    /// </summary>
    IFormatProvider? Provider { get; set; }
}
