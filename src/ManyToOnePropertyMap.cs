using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper
{
    public delegate IEnumerable<T> CreateElementsFactory<T>(IEnumerable<T> elements);

    /// <summary>
    /// Reads multiple cells of an excel sheet and maps the value of the cell to the
    /// type of the property or field.
    /// </summary>
    public abstract class ManyToOnePropertyMap<T> : ExcelPropertyMap
    {
        public IMultipleCellValuesReader _cellValuesReader;

        public IMultipleCellValuesReader CellValuesReader
        {
            get => _cellValuesReader;
            set => _cellValuesReader = value ?? throw new ArgumentNullException(nameof(value));
        }

        public bool Optional { get; set; }

        /// <summary>
        /// Constructs a map that reads one or more values from one or more cells and maps these values to one
        /// property and field of the type of the property or field.
        /// </summary>
        /// <param name="member">The property or field to map the value of a one or more cells to.</param>
        /// <param name="cellValuesReader">The reader.</param>
        public ManyToOnePropertyMap(MemberInfo member, IMultipleCellValuesReader cellValuesReader) : base(member)
        {
            CellValuesReader = cellValuesReader ?? throw new ArgumentNullException(nameof(cellValuesReader));
        }
    }
}
