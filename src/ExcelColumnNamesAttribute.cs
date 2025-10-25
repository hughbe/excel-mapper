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
    private IReadOnlyList<string> _columnNames;
    private StringComparison _comparison = StringComparison.OrdinalIgnoreCase;

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnNamesAttribute"/> with the specified column name.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <param name="comparison">The string comparison to use when matching column names.</param>
    public ExcelColumnNamesAttribute(params string[] columnNames)
    {
        ColumnUtilities.ValidateColumnNames(columnNames);
        _columnNames = columnNames;
    }

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnNamesAttribute"/> with the specified column name.
    /// </summary>
    /// <param name="comparison">The string comparison to use when matching column names.</param>
    /// <param name="columnNames">The names of the columns.</param>
    public ExcelColumnNamesAttribute(string[] columnNames, StringComparison comparison)
    {
        ColumnUtilities.ValidateColumnNames(columnNames);
        EnumUtilities.ValidateIsDefined(comparison);
        _columnNames = columnNames;
        _comparison = comparison;
    }

    /// <summary>
    /// Gets or sets the names of the columns.
    /// </summary>
    public IReadOnlyList<string> Names
    {
        get => _columnNames;
        set
        {
            ColumnUtilities.ValidateColumnNames(value);
            _columnNames = value;
        }
    }

    /// <summary>
    /// Gets or sets the string comparison to use when matching column names.
    /// </summary>
    /// <remarks>The default is <see cref="StringComparison.OrdinalIgnoreCase"/>.</remarks>
    public StringComparison Comparison
    {
        get => _comparison;
        set
        {
            EnumUtilities.ValidateIsDefined(value);
            _comparison = value;
        }
    }
}
