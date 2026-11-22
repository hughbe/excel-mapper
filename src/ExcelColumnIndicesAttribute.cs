// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace ExcelMapper;

/// <summary>
/// Specifies the column indices that are used when deserializing a property or field.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelColumnIndicesAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnIndicesAttribute"/> with the specified column name.
    /// </summary>
    /// <param name="columnIndices">The indices of the columns.</param>
    public ExcelColumnIndicesAttribute(params int[] columnIndices)
    {
        ColumnUtilities.ValidateColumnIndices(columnIndices , nameof(columnIndices));
        Indices = columnIndices;
    }

    /// <summary>
    /// Gets the indices of the columns.
    /// </summary>
    public int[] Indices { get; }
}
