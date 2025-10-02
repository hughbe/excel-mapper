// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;

namespace ExcelMapper;

/// <summary>
/// Prevents a property from being deserialized.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelOptionalAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelOptionalAttribute"/>.
    /// </summary>
    public ExcelOptionalAttribute()
    {
    }
}
