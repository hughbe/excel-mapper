using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Binding;

/// <summary>
/// Reads rows with a frozen <see cref="ExcelClassMap{T}"/>: it creates each
/// <typeparamref name="T"/> as its <see cref="ConstructionPlan{T}"/> says, and sets it with the
/// class map's members bound to the sheet (see <see cref="SheetView.GetBoundMembers{T}"/>), or
/// reads it as a single value for <c>Map(p => p)</c>. Values of a value type aren't boxed.
/// </summary>
internal sealed class ClassMapReader<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap)
{
    // How T is created.
    private readonly ConstructionPlan<T>? _constructionPlan = classMap.SingleValueMap is null ? ConstructionPlan<T>.Create(classMap.Members) : null;

    /// <summary>
    /// Maps a row (or for a nested object, the cells it reads) to a <typeparamref name="T"/>, which
    /// isn't boxed if it's a value type. When <paramref name="errors"/> isn't null (see
    /// <see cref="ExcelSheet.ReadRowResults{T}()"/>), mapping errors are added to it rather than
    /// thrown, and every member is still read.
    /// </summary>
    /// <returns>Whether there is a value: false if a nested object's cells are missing and it's
    /// optional, or errors are being collected and the row couldn't be created.</returns>
    public bool TryRead(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, out T result)
    {
        // If the map represents a single value, use that map to get the value.
        if (classMap.SingleValueMap is { } valueMap)
        {
            if (errors is null)
            {
                return TryGetSingleValue(valueMap, sheet, rowIndex, reader, member, null, out result);
            }

            try
            {
                return TryGetSingleValue(valueMap, sheet, rowIndex, reader, member, errors, out result);
            }
            catch (ExcelMappingException exception)
            {
                errors.Add(exception);
                result = default!;
                return false;
            }
        }

        var plan = _constructionPlan!;
        var members = sheet.GetBoundMembers(classMap);
        if (plan.Constructor is null)
        {
            // Create the instance and set its members.
            result = plan.CreateInstance!();
            MapBoundMembers(members, ref result, sheet, rowIndex, reader, errors);
            return true;
        }

        return TryCreateUsingConstructor(plan, members, sheet, rowIndex, reader, errors, out result);
    }

    private static bool TryGetSingleValue(IMap valueMap, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, out T result)
    {
        // The value map of Map(p => p) reads a T, without boxing.
        if (valueMap is ValueMap<T> typedMap)
        {
            return typedMap.TryGetValue(typedMap.FindColumn(sheet), sheet, rowIndex, reader, member, errors, out result);
        }

        if (valueMap.TryGetValue(sheet, rowIndex, reader, member, errors, out var value))
        {
            result = (T)value;
            return true;
        }

        result = default!;
        return false;
    }

    private static void MapBoundMembers(BoundMember<T>[] members, ref T instance, SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors)
    {
        if (errors is null)
        {
            for (int i = 0; i < members.Length; i++)
            {
                members[i].TryMap(ref instance, sheet, rowIndex, reader, null);
            }

            return;
        }

        for (int i = 0; i < members.Length; i++)
        {
            try
            {
                members[i].TryMap(ref instance, sheet, rowIndex, reader, errors);
            }
            catch (ExcelMappingException exception)
            {
                errors.Add(exception);
            }
        }
    }

    /// <summary>
    /// Reads a member's value for the row, for a type created through a constructor (see
    /// <see cref="BoundMember{TOwner}.ReadValue"/>). When errors are being collected (see
    /// <see cref="ExcelSheet.ReadRowResults{T}()"/>), a mapping error is recorded and the member has no
    /// value, so that the remaining members are still read.
    /// </summary>
    private static void ReadMemberValue(BoundMember<T> member, SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors)
    {
        if (errors is null)
        {
            member.ReadValue(sheet, rowIndex, reader, null);
            return;
        }

        try
        {
            member.ReadValue(sheet, rowIndex, reader, errors);
        }
        catch (ExcelMappingException exception)
        {
            errors.Add(exception);
        }
    }

    private static bool TryCreateUsingConstructor(ConstructionPlan<T> plan, BoundMember<T>[] members, SheetView sheet, int rowIndex, IExcelDataReader reader, ErrorTarget? errors, out T result)
    {
        var errorCount = errors?.Count ?? 0;

        // Read every member first, as constructor arguments must be known before the object exists.
        // Each member keeps its value until it's passed to the constructor or set.
        for (int i = 0; i < members.Length; i++)
        {
            ReadMemberValue(members[i], sheet, rowIndex, reader, errors);
        }

        // When collecting errors, don't call the constructor with arguments that couldn't be read.
        if (errors is not null && errors.Count > errorCount)
        {
            result = default!;
            return false;
        }

        try
        {
            result = plan.Construct(members);
        }
        catch (Exception ex)
        {
            var inner = ex is TargetInvocationException { InnerException: { } invocationInner } ? invocationInner : ex;
            // The constructor's message follows the row, as a hint follows a cell's address.
            var exception = new ExcelMappingException($"Could not create an instance of type {TypeNames.Get(typeof(T))}", sheet, rowIndex, -1, inner, " " + inner.Message);
            if (errors is null)
            {
                throw exception;
            }

            errors.Add(exception);
            result = default!;
            return false;
        }

        plan.SetOtherMembers(ref result, members);
        return true;
    }
}
