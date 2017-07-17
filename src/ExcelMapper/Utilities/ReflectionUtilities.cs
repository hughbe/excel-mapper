using System;
using System.Collections.Generic;
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

        public static bool ImplementsGenericInterface(this Type type, Type genericInterfaceType, out Type elementType)
        {
            foreach (Type interfaceType in type.GetTypeInfo().ImplementedInterfaces)
            {
                if (interfaceType.GetTypeInfo().IsGenericType && interfaceType.GetGenericTypeDefinition() == genericInterfaceType)
                {
                    elementType = interfaceType.GenericTypeArguments[0];
                    return true;
                }
            }

            elementType = null;
            return false;
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
            else if (member is FieldInfo field)
            {
                return field.FieldType;
            }

            throw new ArgumentException($"Member \"{member.Name}\" is not a property or field.", nameof(member));
        }

        public static Type GetNullableTypeOrThis(this Type type, out bool isNullable)
        {
            isNullable = type.IsNullable();
            return isNullable ? type.GenericTypeArguments[0] : type;
        }

        public static bool IsNullable(this Type type)
        {
            return type.GetTypeInfo().IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        public static bool GetElementTypeOrEnumerableType(this Type type, out Type elementType)
        {
            if (type.IsArray)
            {
                elementType = type.GetElementType();
                return true;
            }

            return type.ImplementsGenericInterface(typeof(IEnumerable<>), out elementType);
        }
    }
}
