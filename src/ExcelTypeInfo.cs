using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper;
 
public class ExcelClassMap : IMap
{
    public ExcelClassMap(Type type)
    {
        Type = type ?? throw new ArgumentNullException(nameof(type));
    }

    public Type Type { get; }

    public ExcelPropertyMapCollection Properties { get; } = new ExcelPropertyMapCollection();

    public bool TryGetValue(ExcelRow row, IExcelDataReader reader, MemberInfo member, out object value)
    {
        object instance = Activator.CreateInstance(Type);
        foreach (ExcelPropertyMap property in Properties)
        {
            if (property.Map.TryGetValue(row, reader, property.Member, out object propertyValue))
            {
                property.SetValueFactory(instance, propertyValue);
            }
        }

        value = instance;
        return true;
    }
}
