using System;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    public delegate void SetPropertyDelegate(object instance, object value);

    public abstract class ExcelPropertyMap
    {
        public MemberInfo Member { get; }
        public SetPropertyDelegate SetPropertyFactory { get; }

        public ExcelPropertyMap(MemberInfo member)
        {
            if (member == null)
            {
                throw new ArgumentNullException(nameof(member));
            }

            if (member is PropertyInfo property)
            {
                if (!property.CanWrite)
                {
                    throw new ArgumentException($"Property \"{member.Name}\" is read-only.", nameof(member));
                }

                Member = member;
                SetPropertyFactory = (instance, value) => property.SetValue(instance, value);
            }
            else if (member is FieldInfo field)
            {
                Member = member;
                SetPropertyFactory = (instance, value) => field.SetValue(instance, value);
            }
            else
            {
                throw new ArgumentException($"Member \"{member.Name}\" is not a field or property.", nameof(member));
            }
        }

        public abstract object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}
