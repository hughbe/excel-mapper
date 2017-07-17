namespace ExcelMapper.Mappings
{
    /// <summary>
    /// An enumeration describing the result of an operation to map the value
    /// of a cell to a property or field.
    /// </summary>
    public enum PropertyMappingResultType
    {
        /// <summary>
        /// The value could not be mapped, but is not invalid. This can be used
        /// for optional value mappers.
        /// </summary>
        Continue,

        /// <summary>
        /// The value could be mapped, but prefer the result of mapping items further on in
        /// the mapping pipeline. This can be used for specifiying value mappers that are lower priority.
        /// </summary>
        SuccessIfNoOtherSuccess,

        /// <summary>
        /// The value could be mapped. This mapped value will be used to set the value of the property or
        /// field.
        /// </summary>
        Success,

        /// <summary>
        /// The value was invalid. The InvalidFallback will be invoked if no other value mappers are
        /// successful.
        /// </summary>
        Invalid
    }
}
