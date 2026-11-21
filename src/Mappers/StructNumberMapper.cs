#if NET7_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using System.Numerics;

namespace ExcelMapper.Mappers;

/// <summary>
/// An <see cref="INumberBaseMapper{T}"/> for a number value type, which also converts without boxing
/// for members of its nullable type.
/// </summary>
internal sealed class StructNumberMapper<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)] T> : INumberBaseMapper<T>, ITypedCellMapper<T?> where T : struct, INumberBase<T>
{
    CellConversion ITypedCellMapper<T?>.Convert(ReadCellResult cell, out T? value)
    {
        var conversion = ((ITypedCellMapper<T>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }
}
#endif
