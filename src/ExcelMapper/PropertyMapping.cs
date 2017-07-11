using System;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    public delegate void SetPropertyDelegate(object instance, object value);

    public abstract class PropertyMapping
    {
        public MemberInfo Member { get; }
        public SetPropertyDelegate SetPropertyFactory { get; }

        public PropertyMapping(MemberInfo member)
        {
            if (member == null)
            {
                throw new ArgumentNullException(nameof(member));
            }

            if (member is PropertyInfo property)
            {
                SetPropertyFactory = (instance, value) => property.SetValue(instance, value);
            }
            else if (member is FieldInfo field)
            {
                SetPropertyFactory = (instance, value) => field.SetValue(instance, value);
            }
            else
            {
                throw new ExcelMappingException($"Member {member.Name} is not a field or property.");
            }
        }

        public abstract object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}
