using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Utilities
{
    internal static class ReflectionUtilities
    {
        public static bool ImplementsInterface(this Type type, Type interfaceType)
        {
            return type.GetTypeInfo().ImplementedInterfaces.Any(t => t == interfaceType);
        }

        public static object DefaultValue(this Type type)
        {
            if (type.GetTypeInfo().IsValueType)
            {
                return Activator.CreateInstance(type);
            }

            return null;
        }

        public static Type MemberType(this MemberInfo member)
        {
            if (member is PropertyInfo property)
            {
                return property.PropertyType;
            }

            Debug.Assert(member is FieldInfo);
            return ((FieldInfo)member).FieldType;
        }

        public static bool IsNullable(this Type type)
        {
            return type.GetTypeInfo().IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }
    }
}
