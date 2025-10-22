// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Text.RegularExpressions;
using ExcelMapper.Readers;

namespace ExcelMapper;

/// <summary>
/// Specifies the column matcher that is used when deserializing a property or field.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelColumnsMatchingAttribute : Attribute
{
    /// <summary>
    /// The type of the <see cref="IExcelColumnMatcher"/>.
    /// </summary>
    public Type Type { get; }

    /// <summary>
    /// The constructor arguments for the <see cref="IExcelColumnMatcher"/>.
    /// </summary>
    public object?[]? ConstructorArguments { get; set; }

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnsMatchingAttribute"/> with the specified matcher.
    /// </summary>
    /// <param name="matcherType">The type of the <see cref="IExcelColumnMatcher"/>.</param>
    public ExcelColumnsMatchingAttribute(Type matcherType)
    {
        ArgumentNullException.ThrowIfNull(matcherType);
        if (matcherType.IsAbstract || matcherType.IsInterface)
        {
            throw new ArgumentException("Matcher type cannot be abstract or an interface", nameof(matcherType));
        }
        if (!matcherType.ImplementsInterface(typeof(IExcelColumnMatcher)))
        {
            throw new ArgumentException($"Matcher type must implement {nameof(IExcelColumnMatcher)}", nameof(matcherType));
        }

        Type = matcherType;
    }
    
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnsMatchingAttribute"/> with the specified regex matcher.
    /// </summary>
    /// <param name="regex">The regular expression pattern to match.</param>
    /// <param name="options">A bitwise combination of the enumeration values that modify the regular expression.</param>
    public ExcelColumnsMatchingAttribute(string pattern, RegexOptions options = RegexOptions.None)
    {
        ArgumentNullException.ThrowIfNull(pattern);
        ArgumentException.ThrowIfNullOrEmpty(pattern);

        Type = typeof(RegexColumnMatcher);
        ConstructorArguments = [new Regex(pattern, options)];
    }
}
