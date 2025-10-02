// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;

namespace ExcelMapper;

/// <summary>
/// Provides a value to use if the property or field is empty in Excel.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelDefaultValueAttribute : Attribute
{
    public object? Value { get; }

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelIgnoreAttribute"/>.
    /// </summary>
    public ExcelDefaultValueAttribute(object? value)
    {
        Value = value;
    }
}
