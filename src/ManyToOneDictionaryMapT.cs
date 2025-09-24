using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper
{
    public delegate IDictionary<string, T> CreateDictionaryFactory<T>(IEnumerable<KeyValuePair<string, T>> elements);

    /// <summary>
    /// A map that reads one or more values from one or more cells and maps these values to the type of the
    /// property or field. This is used to map IDictionary properties and fields.
    /// </summary>
    /// <typeparam name="T">The element type of the IDictionary property or field.</typeparam>
    public class ManyToOneDictionaryMap<T> : IMap
    {
        /// <summary>
        /// Constructs a map reads one or more values from one or more cells and maps these values as element
        /// contained by the property or field.
        /// </summary>
        /// <param name="valuePipeline">The map that maps the value of a single cell to an object of the element type of the property or field.</param>
        public ManyToOneDictionaryMap(IMultipleCellValuesReader cellValuesReader, IValuePipeline<T> valuePipeline, CreateDictionaryFactory<T> createDictionaryFactory)
        {
            _cellValuesReader = cellValuesReader ?? throw new ArgumentNullException(nameof(cellValuesReader));
            ValuePipeline = valuePipeline ?? throw new ArgumentNullException(nameof(valuePipeline));
            CreateDictionaryFactory = createDictionaryFactory ?? throw new ArgumentNullException(nameof(createDictionaryFactory));
        }

        /// <summary>
        /// Gets the map that maps the value of a single cell to an object of the element type of the property
        /// or field.
        /// </summary>
        public IValuePipeline<T> ValuePipeline { get; private set; }

        /// <summary>
        /// Gets the reader that reads one or more values from one or more cells used to map each
        /// element of the property or field.
        /// </summary>
        private IMultipleCellValuesReader _cellValuesReader;

        public IMultipleCellValuesReader CellValuesReader
        {
            get => _cellValuesReader;
            set => _cellValuesReader = value ?? throw new ArgumentNullException(nameof(value));
        }

        public CreateDictionaryFactory<T> CreateDictionaryFactory { get; }

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }
            if (sheet.Heading == null)
            {
                throw new ExcelMappingException("The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
            }
            if (!CellValuesReader.TryGetValues(sheet, rowIndex, reader, out IEnumerable<ReadCellValueResult>? valueResults))
            {
                throw new ExcelMappingException($"Could not read value for \"{member?.Name}\"", sheet, rowIndex, -1);
            }

            var valueResultsList = valueResults.ToList();

            var values = new List<T>();
            foreach (ReadCellValueResult valueResult in valueResultsList)
            {
                // Discarding nullability check because it may be indended to be this way (T may be nullable)
                T keyValue = (T)ExcelMapper.ValuePipeline.GetPropertyValue(ValuePipeline, sheet, rowIndex, valueResult, member)!;
                values.Add(keyValue);
            }

            var heading = sheet.Heading;
            IEnumerable<string> keys = valueResultsList.Select(r => heading.GetColumnName(r.ColumnIndex));
            IEnumerable<KeyValuePair<string, T>> elements = keys.Zip(values, (key, keyValue) => new KeyValuePair<string, T>(key, keyValue));
            value = CreateDictionaryFactory(elements);
            return true;
        }

        /// <summary>
        /// Sets the reader of the property map to read the values of one or more cells contained
        /// in the columns with the given names.
        /// </summary>
        /// <param name="columnNames">The name of each column to read.</param>
        /// <returns>The property map that invoked this method.</returns>
        public ManyToOneDictionaryMap<T> WithColumnNames(params string[] columnNames)
        {
            CellValuesReader = new MultipleColumnNamesValueReader(columnNames);
            return this;
        }

        /// <summary>
        /// Sets the reader of the property map to read the values of one or more cells contained
        /// in the columns with the given names.
        /// </summary>
        /// <param name="columnNames">The name of each column to read.</param>
        /// <returns>The property map that invoked this method.</returns>
        public ManyToOneDictionaryMap<T> WithColumnNames(IEnumerable<string> columnNames)
        {
            if (columnNames == null)
            {
                throw new ArgumentNullException(nameof(columnNames));
            }

            return WithColumnNames([.. columnNames]);
        }

        /// <summary>
        /// Sets the map that maps the value of a single cell to an object of the element type of the property
        /// or field.
        /// </summary>
        /// <param name="valueMap">The pipeline that maps the value of a single cell to an object of the element type of the property
        /// or field.</param>
        /// <returns>The property map that invoked this method.</returns>
        public ManyToOneDictionaryMap<T> WithValueMap(Func<IValuePipeline<T>, IValuePipeline<T>> valueMap)
        {
            if (valueMap == null)
            {
                throw new ArgumentNullException(nameof(valueMap));
            }

            ValuePipeline = valueMap(ValuePipeline) ?? throw new ArgumentNullException(nameof(valueMap));
            return this;
        }
    }
}
