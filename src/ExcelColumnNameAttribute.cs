// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;

namespace ExcelMapper
{
    /// <summary>
    /// Specifies the column name that is used when deserializing a property
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public sealed class ExcelColumnNameAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of <see cref="ExcelColumnNameAttribute"/> with the specified column name.
        /// </summary>
        /// <param name="name">The name of the column.</param>
        public ExcelColumnNameAttribute(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            if (name.Length == 0)
            {
                throw new ArgumentException("Column name cannot be empty.", nameof(name));
            }

            Name = name;
        }

        /// <summary>
        /// The name of the column.
        /// </summary>
        public string Name { get; }
    }
}
