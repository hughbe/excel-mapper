// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;

namespace ExcelMapper
{
    /// <summary>
    /// Specifies the column index that is used when deserializing a property
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public sealed class ExcelColumnIndexAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of <see cref="ExcelColumnIndexAttribute"/> with the specified column index.
        /// </summary>
        /// <param name="index">The index of the column.</param>
        public ExcelColumnIndexAttribute(int index)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index), index, $"Column index {index} must be greater or equal to zero.");
            }

            Index = index;
        }

        /// <summary>
        /// The index of the column.
        /// </summary>
        public int Index { get; }
    }
}
