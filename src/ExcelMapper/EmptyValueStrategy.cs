namespace ExcelMapper
{
    /// <summary>
    /// Describes the strategy used to fallback when the value of a cell is empty or invalid.
    /// </summary>
    public enum FallbackStrategy
    {
        /// <summary>
        /// Throw an ExcelMappingException if the value of a cell is empty or invalid.
        /// </summary>
        ThrowIfPrimitive,

        /// <summary>
        /// Return the default value of the type of a property or field if the value of a
        /// cell is empty or invalid.
        /// </summary>
        SetToDefaultValue
    }
}
