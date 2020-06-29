using System;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    public class ExcelClassMap : IMap
    {
        public ExcelClassMap(Type type)
        {
            Type = type ?? throw new ArgumentNullException(nameof(type));
        }

        public Type Type { get; }

        public ExcelPropertyMapCollection Properties { get; } = new ExcelPropertyMapCollection();

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo member, out object result)
        {
            object instance = Activator.CreateInstance(Type);
            foreach (ExcelPropertyMap property in Properties)
            {
                if (property.Map.TryGetValue(sheet, rowIndex, reader, property.Member, out object value))
                {
                    property.SetValueFactory(instance, value);
                }
            }

            result = instance;
            return true;
        }
    }
}
