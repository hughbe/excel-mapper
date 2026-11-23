#if NET7_0_OR_GREATER
using System.Numerics;

namespace ExcelMapper.Mappers;

/// <summary>
/// An <see cref="INumberBaseMapper{T}"/> for a number value type, which also converts without boxing
/// for members of its nullable type.
/// </summary>
internal sealed class StructNumberMapper<T> : INumberBaseMapper<T>, ICellMapper<T?> where T : struct, INumberBase<T>
{
    CellConversion ICellMapper<T?>.Convert(ReadCellResult cell, in CellContext context, out T? value, ref object? error)
    {
        var conversion = ((ICellMapper<T>)this).Convert(cell, in context, out var result, ref error);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }
}
#endif
