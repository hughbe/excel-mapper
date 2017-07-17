using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Readers;
using ExcelMapper.Mappings.Support;

namespace ExcelMapper
{
    /// <summary>
    /// Reads a single cell of an excel sheet and maps the value of the cell to the
    /// type of the property or field.
    /// </summary>
    public class SingleExcelPropertyMap : ExcelPropertyMap, ISinglePropertyMapping
    {
        private readonly List<ICellValueTransformer> _cellValueTransformers = new List<ICellValueTransformer>();
        private readonly List<ICellValueMapper> _cellValueMappers = new List<ICellValueMapper>();

        /// <summary>
        /// Gets or sets the object that takes a sheet and row index and produces the value of a cell.
        /// TODO: check for null. 
        /// </summary>
        public ICellValueReader CellReader { get; set; }

        /// <summary>
        /// Gets the list of objects that take the initial string value read from a cell and
        /// modifies the string value. This is useful for things like trimming the string value
        /// before mapping it.
        /// </summary>
        public IEnumerable<ICellValueTransformer> CellValueTransformers => _cellValueTransformers;

        /// <summary>
        /// Gets the pipeline of items that take the initial string value read from a cell and
        /// converts the string value into the type of the property or field. The items form
        /// a pipeline: if a mapper fails to parse or map the cell value, the next item is used.
        /// </summary>
        public IEnumerable<ICellValueMapper> CellValueMappers => _cellValueMappers;

        /// <summary>
        /// Adds the given mapper to the pipeline of cell value mappers.        
        /// </summary>
        /// <param name="mapper">The mapper to add.</param>
        public void AddCellValueMapper(ICellValueMapper mapper)
        {
            if (mapper == null)
            {
                throw new ArgumentNullException(nameof(mapper));
            }

            _cellValueMappers.Add(mapper);
        }

        /// <summary>
        /// Removes the mapper at the given index from the pipeline of cell value mappers.
        /// </summary>
        /// <param name="index">The index of the mapper to remove.</param>
        public void RemoveCellValueMapper(int index) => _cellValueMappers.RemoveAt(index);

        /// <summary>
        /// Adds the given transformer to the pipeline of cell value transformers.
        /// </summary>
        /// <param name="transformer">The tranformer to add.</param>
        public void AddCellValueTransformer(ICellValueTransformer transformer)
        {
            if (transformer == null)
            {
                throw new ArgumentNullException(nameof(transformer));
            }

            _cellValueTransformers.Add(transformer);
        }

        /// <summary>
        /// Gets or sets an object that handles mapping a cell value to a property or field if the value of the
        /// cell is empty. For example, you can provide a fixed value to return if the value of the cell
        /// is empty.
        /// </summary>
        public IFallbackItem EmptyFallback { get; set; }

        /// <summary>
        /// Gets or sets an object that handles mapping a cell value to a property or field if all items
        /// in the mapper pipeline failed to map the value to the property or field. For example, you can
        /// provide a fixed value to return if the value of the cell is invalid.
        /// </summary>
        public IFallbackItem InvalidFallback { get; set; }

        /// <summary>
        /// Constructs a map that reads the value of a single cell and maps the value of the cell
        /// to the type of the property or field.
        /// </summary>
        /// <param name="member">The property or field to map the value of a single cell to.</param>
        public SingleExcelPropertyMap(MemberInfo member) : base(member)
        {
            CellReader = new ColumnNameValueReader(member.Name);
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            ReadCellValueResult readResult = CellReader.GetValue(sheet, rowIndex, reader);
            return GetPropertyValue(sheet, rowIndex, reader, readResult);
        }

        internal object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, ReadCellValueResult readResult)
        {
            foreach (ICellValueTransformer transformer in _cellValueTransformers)
            {
                readResult = new ReadCellValueResult(readResult.ColumnIndex, transformer.TransformStringValue(sheet, rowIndex, readResult));
            }

            if (readResult.StringValue == null && EmptyFallback != null)
            {
                return EmptyFallback.PerformFallback(sheet, rowIndex, readResult);
            }

            PropertyMappingResultType resultType = PropertyMappingResultType.Success;
            object value = null;

            foreach (ICellValueMapper mappingItem in _cellValueMappers)
            {
                PropertyMappingResultType newResultType  = mappingItem.GetProperty(readResult, ref value);
                if (newResultType == PropertyMappingResultType.Success)
                {
                    return value;
                }

                if (newResultType != PropertyMappingResultType.Continue)
                {
                    resultType = newResultType;
                }
            }

            if (resultType != PropertyMappingResultType.Success && resultType != PropertyMappingResultType.SuccessIfNoOtherSuccess && InvalidFallback != null)
            {
                return InvalidFallback.PerformFallback(sheet, rowIndex, readResult);
            }

            return value;
        }
    }
}
