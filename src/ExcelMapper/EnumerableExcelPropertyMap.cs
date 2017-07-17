using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Readers;

namespace ExcelMapper
{
    /// <summary>
    /// A map that reads one or more values from one or more cells and maps these values to the type of the
    /// property or field. This is used to map IEnumerable properties and fields.
    /// </summary>
    /// <typeparam name="T">The element type of the IEnumerable property or field.</typeparam>
    public abstract class EnumerableExcelPropertyMap<T> : ExcelPropertyMap
    {
        /// <summary>
        /// Gets the map that maps the value of a single cell to an object of the element type of the property
        /// or field.
        /// </summary>
        public SingleExcelPropertyMap<T> ElementMap { get; private set; }

        /// <summary>
        /// Gets the reader that reads one or more values from one or more cells used to map each
        /// element of the property or field.
        /// </summary>
        public IMultipleCellValuesReader ColumnsReader { get; private set; }

        /// <summary>
        /// Constructs a map reads one or more values from one or more cells and maps these values as element
        /// contained by the property or field.
        /// </summary>
        /// <param name="member">The property or field to map the values of one or more cell to.</param>
        /// <param name="elementMapping">The map that maps the value of a single cell to an object of the element type of the property or field.</param>
        protected EnumerableExcelPropertyMap(MemberInfo member, SingleExcelPropertyMap<T> elementMapping) : base(member)
        {
            ElementMap = elementMapping ?? throw new ArgumentNullException(nameof(elementMapping));

            var columnReader = new ColumnNameValueReader(member.Name);
            ColumnsReader = new SplitCellValueReader(columnReader); 
        }

        /// <summary>
        /// Sets the map that maps the value of a single cell to an object of the element type of the property
        /// or field.
        /// </summary>
        /// <param name="elementMap">The map that maps the value of a single cell to an object of the element type of the property
        /// or field.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithElementMap(Func<SingleExcelPropertyMap<T>, SingleExcelPropertyMap<T>> elementMap)
        {
            if (elementMap == null)
            {
                throw new ArgumentNullException(nameof(elementMap));
            }

            ElementMap = elementMap(ElementMap) ?? throw new ArgumentNullException(nameof(elementMap));
            return this;
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            ReadCellValueResult[] values = ColumnsReader.GetValues(sheet, rowIndex, reader).ToArray();
            var elements = new List<T>(values.Length);

            foreach (ReadCellValueResult value in values)
            {
                T elementValue = (T)ElementMap.GetPropertyValue(sheet, rowIndex, reader, value);
                elements.Add(elementValue);
            }

            return CreateFromElements(elements);
        }

        /// <summary>
        /// Creates an object that implements IEnumerable&lt;T&gt; given a list of elements.
        /// </summary>
        /// <param name="elements">The elements created from the mapped values of one or more cells.</param>
        /// <returns>An object that will be used to assign the value of the property or field.</returns>
        protected abstract object CreateFromElements(IEnumerable<T> elements);

        /// <summary>
        /// Sets the reader for multiple values to split the value of a single cell contained in the column
        /// with a given name.
        /// </summary>
        /// <param name="columnName">The name of the column containing the cell to split.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithColumnName(string columnName)
        {
            var columnReader = new ColumnNameValueReader(columnName);
            if (ColumnsReader is SplitCellValueReader splitColumnReader)
            {
                splitColumnReader.CellReader = columnReader;
            }
            else
            {
                ColumnsReader = new SplitCellValueReader(columnReader);
            }

            return this;
        }

        /// <summary>
        /// Sets the reader for multiple values to split the value of a single cell contained in the column
        /// at the given zero-based index.
        /// </summary>
        /// <param name="columnIndex">The zero-bassed index of the column containing the cell to split.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithColumnIndex(int columnIndex)
        {
            var reader = new ColumnIndexValueReader(columnIndex);
            if (ColumnsReader is SplitCellValueReader splitColumnReader)
            {
                splitColumnReader.CellReader = reader;
            }
            else
            {
                ColumnsReader = new SplitCellValueReader(reader);
            }

            return this;
        }

        /// <summary>
        /// Sets the reader of the property map to split the value of a single cell using the
        /// given separators.
        /// </summary>
        /// <param name="separators">The separators used to split the value of a single cell.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithSeparators(params char[] separators)
        {
            if (!(ColumnsReader is SplitCellValueReader splitColumnReader))
            {
                throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
            }

            splitColumnReader.Separators = separators;
            return this;
        }

        /// <summary>
        /// Sets the reader of the property map to split the value of a single cell using the
        /// given separators.
        /// </summary>
        /// <param name="separators">The separators used to split the value of a single cell.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithSeparators(IEnumerable<char> separators)
        {
            return WithSeparators(separators?.ToArray());
        }

        /// <summary>
        /// Sets the reader of the property map to read the values of one or more cells contained
        /// in the columns with the given names.
        /// </summary>
        /// <param name="columnNames">The name of each column to read.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithColumnNames(params string[] columnNames)
        {
            ColumnsReader = new MultipleColumnNamesValueReader(columnNames);
            return this;
        }

        /// <summary>
        /// Sets the reader of the property map to read the values of one or more cells contained
        /// in the columns with the given names.
        /// </summary>
        /// <param name="columnNames">The name of each column to read.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithColumnNames(IEnumerable<string> columnNames)
        {
            return WithColumnNames(columnNames?.ToArray());
        }

        /// <summary>
        /// Sets the reader of the property map to read the values of one or more cells contained
        /// in the columns with the given zero-based indices.
        /// </summary>
        /// <param name="columnIndices">The zero-based index of each column to read.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithColumnIndices(params int[] columnIndices)
        {
            ColumnsReader = new MultipleColumnIndicesValueReader(columnIndices);
            return this;
        }

        /// <summary>
        /// Sets the reader of the property map to read the values of one or more cells contained
        /// in the columns with the given zero-based indices.
        /// </summary>
        /// <param name="columnIndices">The zero-based index of each column to read.</param>
        /// <returns>The property map that invoked this method.</returns>
        public EnumerableExcelPropertyMap<T> WithColumnIndices(IEnumerable<int> columnIndices)
        {
            return WithColumnIndices(columnIndices?.ToArray());
        }
    }
}
