// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace ExcelMapper;

/// <summary>
/// Specifies the column names that are used when deserializing a property or field.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelColumnNamesAttribute : Attribute
{
    private string[] _columnNames;

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnNamesAttribute"/> with the specified column name.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    public ExcelColumnNamesAttribute(params string[] columnNames)
    {
        ColumnUtilities.ValidateColumnNames(columnNames , nameof(columnNames));
        _columnNames = columnNames;
    }

    /// <summary>
    /// The names of the columns.
    /// </summary>
    public string[] Names
    {
        get => _columnNames;
        set
        {
            ColumnUtilities.ValidateColumnNames(value , nameof(value));
            _columnNames = value;
        }
    }
}
