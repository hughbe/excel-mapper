// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace ExcelMapper;

/// <summary>
/// Maps a value to another using a dictionary.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public sealed class ExcelMappingDictionaryAttribute : Attribute
{
    /// <summary>
    /// Gets the value to be mapped.
    /// </summary>
    public string Value { get; }
    
    /// <summary>
    /// Gets the mapped value.
    /// </summary>
    public object? MappedValue { get; }

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelMappingDictionaryAttribute"/>.
    /// </summary>
    public ExcelMappingDictionaryAttribute(string value, object? mappedValue)
    {
        ArgumentException.ThrowIfNullOrEmpty(value);
        Value = value;
        MappedValue = mappedValue;
    }
}
