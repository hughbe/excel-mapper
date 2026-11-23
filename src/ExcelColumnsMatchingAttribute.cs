using System.Diagnostics.CodeAnalysis;
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
public sealed class ExcelColumnsMatchingAttribute : Attribute
{
    /// <summary>
    /// The type of the <see cref="IColumnSelector"/>.
    /// </summary>
    [DynamicallyAccessedMembers(TrimmingConstants.ActivatedType)]
    public Type Type { get; }

    /// <summary>
    /// The constructor arguments for the <see cref="IColumnSelector"/>.
    /// </summary>
    public object?[]? ConstructorArguments { get; set; }

    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnsMatchingAttribute"/> with the specified matcher.
    /// </summary>
    /// <param name="matcherType">The type of the <see cref="IColumnSelector"/>.</param>
    public ExcelColumnsMatchingAttribute([DynamicallyAccessedMembers(TrimmingConstants.ActivatedType)] Type matcherType)
    {
        ThrowHelpers.ThrowIfNull(matcherType, nameof(matcherType));
        if (matcherType.IsAbstract || matcherType.IsInterface)
        {
            throw new ArgumentException("Matcher type cannot be abstract or an interface", nameof(matcherType));
        }
        if (!matcherType.ImplementsInterface(typeof(IColumnSelector)))
        {
            throw new ArgumentException($"Matcher type must implement {nameof(IColumnSelector)}", nameof(matcherType));
        }

        Type = matcherType;
    }
    
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelColumnsMatchingAttribute"/> that matches column names with a
    /// regular expression, with the options in <see cref="RegexOptions"/>.
    /// </summary>
    /// <param name="pattern">The regular expression pattern to match.</param>
    public ExcelColumnsMatchingAttribute(string pattern)
    {
        ThrowHelpers.ThrowIfNull(pattern, nameof(pattern));
        ThrowHelpers.ThrowIfNullOrEmpty(pattern, nameof(pattern));

        _pattern = pattern;
        Type = typeof(RegexColumnMatcher);
        ConstructorArguments = [new Regex(pattern)];
    }

    // The pattern given to the constructor, if any.
    private readonly string? _pattern;
    private RegexOptions _regexOptions;

    /// <summary>
    /// Gets or sets the options of the regular expression given to the constructor, such as
    /// <c>[ExcelColumnsMatching("^Q\\d$", RegexOptions = RegexOptions.IgnoreCase)]</c>.
    /// </summary>
    /// <exception cref="InvalidOperationException">The attribute was given a selector type rather
    /// than a pattern.</exception>
    public RegexOptions RegexOptions
    {
        get => _regexOptions;
        set
        {
            if (_pattern is null)
            {
                throw new InvalidOperationException("RegexOptions only applies to an attribute given a pattern.");
            }

            // Validated by the Regex constructor.
            ConstructorArguments = [new Regex(_pattern, value)];
            _regexOptions = value;
        }
    }
}
