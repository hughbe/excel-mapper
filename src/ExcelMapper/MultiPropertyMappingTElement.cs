using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    public abstract class MultiPropertyMapping<TElement> : MultiPropertyMapping
    {
        public SinglePropertyMapping<TElement> ElementMapping { get; private set; }

        internal MultiPropertyMapping(MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(member)
        {
            ElementMapping = new SinglePropertyMapping<TElement>(member, emptyValueStrategy);
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            IEnumerable<int> columnIndices = Mapper.GetColumnIndices(sheet, rowIndex, reader);
            var elements = new List<TElement>(Mapper.CapacityEstimate);

            foreach (int columnIndex in columnIndices)
            {
                TElement value = (TElement)ElementMapping.GetPropertyValue(sheet, rowIndex, reader, columnIndex);
                elements.Add(value);
            }

            return CreateFromElements(elements);
        }

        public MultiPropertyMapping<TElement> WithElementMapping(Func<SinglePropertyMapping<TElement>, SinglePropertyMapping<TElement>> elementMapping)
        {
            ElementMapping = elementMapping(ElementMapping);
            return this;
        }

        public abstract object CreateFromElements(IEnumerable<TElement> elements);
    }
}
