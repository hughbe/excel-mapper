#if NET7_0_OR_GREATER
namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper for <see cref="IParsable{TSelf}"/> types, using <c>TParse.Parse</c>.
/// </summary>
internal class ParsableMapper<TParse> : ICellMapper<TParse> where TParse : IParsable<TParse>
{
    public IFormatProvider? Provider { get; set; }

    CellConversion ICellMapper<TParse>.Convert(ReadCellResult cell, out TParse value)
    {
        // TryParse isn't used: types may only implement Parse properly, and Parse's exception says
        // why the value is invalid.
        try
        {
            value = TParse.Parse(cell.GetString()!, Provider ?? cell.Culture);
            return CellConversion.Success;
        }
        catch
        {
            value = default!;
            return CellConversion.Invalid;
        }
    }

    public Exception CreateException(ReadCellResult cell)
        => MapperUtilities.InvalidFromParse(() => TParse.Parse(cell.GetString()!, Provider ?? cell.Culture));
}

/// <summary>
/// A mapper for <see cref="IParsable{TSelf}"/> value types, which also maps nullable members.
/// </summary>
internal sealed class StructParsableMapper<TParse> : ParsableMapper<TParse>, ICellMapper<TParse?> where TParse : struct, IParsable<TParse>
{
    CellConversion ICellMapper<TParse?>.Convert(ReadCellResult cell, out TParse? value)
    {
        var conversion = ((ICellMapper<TParse>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }
}
#endif
