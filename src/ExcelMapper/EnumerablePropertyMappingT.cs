using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Mappings;

namespace ExcelMapper
{
    public abstract class EnumerablePropertyMapping<T> : EnumerablePropertyMapping
    {
        public SinglePropertyMapping<T> ElementMapping { get; private set; }
        public EmptyValueStrategy EmptyValueStrategy { get; }

        public IMultiPropertyMapper Mapper { get; internal set; }

        internal EnumerablePropertyMapping(MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(member)
        {
            ElementMapping = new SinglePropertyMapping<T>(member, emptyValueStrategy);
            EmptyValueStrategy = emptyValueStrategy;

            var columnMapper = new ColumnPropertyMapper(member.Name);
            Mapper = new SplitPropertyMapper(columnMapper); 
        }

        public EnumerablePropertyMapping<T> WithElementMapping(Func<SinglePropertyMapping<T>, SinglePropertyMapping<T>> elementMapping)
        {
            ElementMapping = elementMapping(ElementMapping);
            return this;
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            MapResult[] values = Mapper.GetValues(sheet, rowIndex, reader).ToArray();
            var elements = new List<T>(values.Length);

            for (int i = 0; i < values.Length; i++)
            {
                T value = (T)ElementMapping.GetPropertyValue(sheet, rowIndex, reader, values[i]);
                elements.Add(value);
            }

            return CreateFromElements(elements);
        }

        public abstract object CreateFromElements(IEnumerable<T> elements);
    }
}
