// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace ExcelMapper;

/// <summary>
/// Specifies the column name that is used when deserializing a property or field.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public class ExcelColumnNameAttribute : Attribute
{
    private string _columnName;

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnNameAttribute"/> with the specified column name.
    /// </summary>
    /// <param name="name">The name of the column.</param>
    public ExcelColumnNameAttribute(string columnName)
    {
        ColumnUtilities.ValidateColumnName(columnName , nameof(columnName));
        _columnName = columnName;
    }

    /// <summary>
    /// Gets or sets the name of the column.
    /// </summary>
    public string Name
    {
        get => _columnName;
        set
        {
            ColumnUtilities.ValidateColumnName(value , nameof(value));
            _columnName = value;
        }
    }
}
