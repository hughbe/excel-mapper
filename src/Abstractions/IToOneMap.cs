namespace ExcelMapper.Abstractions;

public interface IToOneMap : IMap
{
    /// <summary>
    /// Gets or sets whether an exception should be thrown if the value cannot be found.
    /// </summary>
    public bool Optional { get; set; }

    /// <summary>
    /// Gets or sets whether the map peserves formatting when reading string values.
    /// </summary>
    public bool PreserveFormatting { get; set; }
}
