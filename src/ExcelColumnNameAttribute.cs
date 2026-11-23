// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace ExcelMapper;

/// <summary>
/// Specifies the column name that is used when deserializing a property or field.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true)]
public sealed class ExcelColumnNameAttribute : Attribute
{
    private StringComparison _comparison = StringComparison.OrdinalIgnoreCase;

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnNameAttribute"/> with the specified column
    /// name, matched ignoring case unless <see cref="Comparison"/> is set.
    /// </summary>
    /// <param name="columnName">The name of the column.</param>
    public ExcelColumnNameAttribute(string columnName)
    {
        ColumnUtilities.ValidateColumnName(columnName);
        Name = columnName;
    }

    /// <summary>
    /// Gets the name of the column.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Gets or sets the string comparison to use when matching the column name, such as
    /// <c>[ExcelColumnName("Name", Comparison = StringComparison.Ordinal)]</c>. The default is
    /// <see cref="StringComparison.OrdinalIgnoreCase"/>.
    /// </summary>
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
