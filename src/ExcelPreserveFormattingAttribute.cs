// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace ExcelMapper;

/// <summary>
/// Reads the value of a cell with formatting.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelPreserveFormattingAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelPreserveFormattingAttribute"/>.
    /// </summary>
    public ExcelPreserveFormattingAttribute()
    {
    }
}
