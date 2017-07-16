using System;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    public class ObjectPropertyMapping<T> : SinglePropertyMapping<T>
    {
        private ExcelClassMap<T> _classMap;

        public ExcelClassMap<T> ClassMap
        {
            get => _classMap;
            set => _classMap = value ?? throw new ArgumentNullException(nameof(value));
        }

        public ObjectPropertyMapping(MemberInfo member, ExcelClassMap<T> classMap) : base(member)
        {
            ClassMap = classMap ?? throw new ArgumentNullException(nameof(classMap));
        }

        public ObjectPropertyMapping<T> WithClassMap(Action<ExcelClassMap<T>> classMapFactory)
        {
            if (classMapFactory == null)
            {
                throw new ArgumentNullException(nameof(classMapFactory));
            }

            classMapFactory(ClassMap);
            return this;
        }

        public ObjectPropertyMapping<T> WithClassMap(ExcelClassMap<T> classMap)
        {
            ClassMap = classMap ?? throw new ArgumentNullException(nameof(classMap));
            return this;
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return _classMap.Execute(sheet, rowIndex, reader);
        }
    }
}
