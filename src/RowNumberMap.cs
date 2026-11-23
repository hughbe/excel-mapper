using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// A map that sets a property or field to the number of the row being read, rather than to
/// the value of a cell. Row numbers are one-based, as Excel shows them, and count every row
/// in the sheet including the heading, so they match the rows in error messages.
/// </summary>
/// <remarks>
/// Create one with <see cref="ExcelRowNumberAttribute"/> or
/// <see cref="ExcelClassMap{T}.MapRowNumber{TProperty}(System.Linq.Expressions.Expression{Func{T, TProperty}})"/>.
/// </remarks>
public sealed class RowNumberMap : IMap
{
    private readonly TypeCode _typeCode;

    /// <summary>
    /// Initializes a new instance of <see cref="RowNumberMap"/> for a member of the given type.
    /// </summary>
    /// <param name="type">The type of the member: a numeric type such as <see cref="int"/> or
    /// <see cref="long"/>, or a nullable numeric type.</param>
    /// <remarks>Callers check the type with <see cref="IsSupportedType"/>.</remarks>
    internal RowNumberMap(Type type)
    {
        Type = type;
        _typeCode = Type.GetTypeCode(type.GetNullableTypeOrThis(out _));
    }

    /// <summary>
    /// Gets the type of the member the row number is assigned to.
    /// </summary>
    public Type Type { get; }

    /// <inheritdoc />
    void IMap.Freeze()
    {
        // A row number map has nothing to change.
    }

    internal static bool IsSupportedType(Type type)
    {
        var underlyingType = type.GetNullableTypeOrThis(out _);
        return !underlyingType.IsEnum && Type.GetTypeCode(underlyingType) is >= TypeCode.SByte and <= TypeCode.Decimal;
    }

    /// <inheritdoc/>
    bool IMap.TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        // Row indexes are zero-based; row numbers are one-based, as in Excel.
        var rowNumber = (long)rowIndex + 1;
        if (_typeCode == TypeCode.Int32)
        {
            // The common case: an int member and a sheet with fewer than 2^31 rows.
            value = (int)rowNumber;
            return true;
        }

        try
        {
            value = _typeCode switch
            {
                TypeCode.SByte => checked((sbyte)rowNumber),
                TypeCode.Byte => checked((byte)rowNumber),
                TypeCode.Int16 => checked((short)rowNumber),
                TypeCode.UInt16 => checked((ushort)rowNumber),
                TypeCode.UInt32 => checked((uint)rowNumber),
                TypeCode.Int64 => rowNumber,
                TypeCode.UInt64 => checked((ulong)rowNumber),
                TypeCode.Single => (float)rowNumber,
                TypeCode.Double => (double)rowNumber,
                _ => (object)(decimal)rowNumber,
            };
            return true;
        }
        catch (OverflowException ex)
        {
            var target = member is null ? $"type \"{Type}\"" : $"member \"{member.Name}\" of type \"{Type}\"";
            throw new ExcelMappingException($"Row number {rowNumber} is too large for {target}", sheet, rowIndex, -1, ex)
            {
                Member = member
            };
        }
    }
}
