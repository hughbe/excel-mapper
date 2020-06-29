using System;
using System.Reflection;

namespace ExcelMapper
{
    public delegate void MemberSetValueDelegate(object instance, object value);

    public class ExcelPropertyMap
    {
        public ExcelPropertyMap(MemberInfo member, IMap map)
        {
            Member = member ?? throw new ArgumentNullException(nameof(member));
            Map = map ?? throw new ArgumentNullException(nameof(map));

            if (member is PropertyInfo property)
            {
                if (!property.CanWrite)
                {
                    throw new ArgumentException($"Property \"{member.Name}\" is read-only.", nameof(member));
                }

                SetValueFactory = (instance, value) => property.SetValue(instance, value);
            }
            else if (member is FieldInfo field)
            {
                SetValueFactory = (instance, value) => field.SetValue(instance, value);
            }
            else
            {
                throw new ArgumentException($"Member \"{member.Name}\" is not a field or property.", nameof(member));
            }
        }

        public MemberInfo Member { get; }

        public MemberSetValueDelegate SetValueFactory { get; }

        public IMap Map { get; }
    }
}
