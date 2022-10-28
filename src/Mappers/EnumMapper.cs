﻿using System.Reflection;
using ExcelMapper.Abstractions;

/// <summary>
/// A mapper that tries to map the value of a cell to an enum of a given type.
/// </summary>
public class EnumMapper : ICellValueMapper
{
    /// <summary>
    /// Gets the type of the enum to map the value of a cell to.
    /// </summary>
    public Type EnumType { get; }

    /// <summary>
    /// Gets whether enum parsing is case insensitive.
    /// </summary>
    public bool IgnoreCase { get; }

    /// <summary>
    /// Constructs a mapper that tries to map the value of a cell to an enum of a given type.
    /// </summary>
    /// <param name="enumType">The type of the enum to convert the value of a cell to.</param>
    /// <param name="ignoreCase">A flag indicating whether enum parsing is case insensitive.</param>
    public EnumMapper(Type enumType) : this(enumType, ignoreCase: false)
    {
    }

    public EnumMapper(Type enumType, bool ignoreCase)
    {
        if (enumType == null)
        {
            throw new ArgumentNullException(nameof(enumType));
        }

        if (!enumType.GetTypeInfo().IsEnum)
        {
            throw new ArgumentException($"Type {enumType} is not an Enum.", nameof(enumType));
        }

        EnumType = enumType;
        IgnoreCase = ignoreCase;
    }

    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        string stringValue = previous.Value?.ToString();

        try
        {
            object result = Enum.Parse(EnumType, stringValue, IgnoreCase);
            return CellValueMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellValueMapperResult.Invalid(exception);
        }
    }
}
