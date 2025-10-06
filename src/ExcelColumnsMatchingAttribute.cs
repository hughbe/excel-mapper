// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Text.RegularExpressions;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using ExcelMapper.Utilities;

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
    /// <param name="name">The name of the column.</param>
    public ExcelColumnsMatchingAttribute(Type matcherType)
    {
        if (matcherType == null)
        {
            throw new ArgumentNullException(nameof(matcherType));
        }
        if (!matcherType.ImplementsInterface(typeof(IExcelColumnMatcher)))
        {
            throw new ArgumentException("Matcher type must implement IExcelColumnMatcher", nameof(matcherType));
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
        if (pattern == null)
        {
            throw new ArgumentNullException(nameof(pattern));
        }
        if (pattern.Length == 0)
        {
            throw new ArgumentException("Pattern cannot be empty", nameof(pattern));
        }

        Type = typeof(RegexColumnMatcher);
        ConstructorArguments = new object[] { new Regex(pattern, options) };
    }
}
