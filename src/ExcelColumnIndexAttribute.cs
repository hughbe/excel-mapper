// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace ExcelMapper;

/// <summary>
/// Specifies the column index that is used when deserializing a property or field.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public class ExcelColumnIndexAttribute : Attribute
{
    private int _index;

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnIndexAttribute"/> with the specified column index.
    /// </summary>
    /// <param name="index">The index of the column.</param>
    public ExcelColumnIndexAttribute(int index)
    {
        ColumnUtilities.ValidateColumnIndex(index, nameof(index));
        Index = index;
    }

    /// <summary>
    /// The index of the column.
    /// </summary>
    public int Index
    {
        get => _index;
        set
        {
            ColumnUtilities.ValidateColumnIndex(value, nameof(value));
            _index = value;
        }
    }
}
