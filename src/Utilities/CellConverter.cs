namespace ExcelMapper.Utilities;

/// <summary>
/// The result of <see cref="ICellConverter{T}.Convert"/>, matching <see cref="CellMapperResult.HandleAction"/>.
/// </summary>
internal enum CellConversion
{
    /// <summary>The value was converted; later converters aren't tried.</summary>
    Success,

    /// <summary>The value was converted, but a later converter's result replaces it.</summary>
    SuccessUnlessReplaced,

    /// <summary>The value is invalid; a later converter's result replaces it.</summary>
    Invalid,

    /// <summary>The converter doesn't handle the value; it's skipped.</summary>
    NotHandled,
}

/// <summary>
/// A mapper that converts a cell to a <typeparamref name="T"/> without boxing. It makes the same
/// decisions as its <see cref="ICellMapper.Map"/>, which is used to create the exception for an
/// invalid value when a fallback needs it.
/// </summary>
internal interface ICellConverter<T> : ICellMapper
{
    CellConversion Convert(ReadCellResult cell, out T value);
}
