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
    public class ExcelColumnNameAttribute : Attribute
    {
        private string _name;

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelColumnNameAttribute"/> with the specified column name.
        /// </summary>
        /// <param name="name">The name of the column.</param>
        public ExcelColumnNameAttribute(string name)
        {
            Name = name;
        }

        /// <summary>
        /// The name of the column.
        /// </summary>
        public string Name
        {
            get => _name;
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                if (value.Length == 0)
                {
                    throw new ArgumentException("Column name cannot be empty.", nameof(value));
                }
                
                _name = value;
            }
        }
    }
}
