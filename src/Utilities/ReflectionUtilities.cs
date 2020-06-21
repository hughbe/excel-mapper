using System;
using System.Collections.Generic;
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
            bool CheckInterface(Type interfaceType, out Type elementTypeResult)
            {
                if (interfaceType.GetTypeInfo().IsGenericType && interfaceType.GetGenericTypeDefinition() == genericInterfaceType)
                {
                    elementTypeResult = interfaceType.GenericTypeArguments[0];
                    return true;
                }

                elementTypeResult = null;
                return false;
            }

            // This type may actually be the interface in question.
            // So return true if this is the case.
            if (type.GetTypeInfo().IsInterface && CheckInterface(type, out elementType))
            {
                return true;
            }

            foreach (Type interfaceType in type.GetTypeInfo().ImplementedInterfaces)
            {
                if (CheckInterface(interfaceType, out elementType))
                {
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

        private static bool IsNullable(this Type type)
        {
            return type.GetTypeInfo().IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        /// <summary>
        /// Gets the element type or the IEnumerable<T> type of the given type.
        /// </summary>
        /// <param name="type">The type to get the element type of.</param>
        /// <param name="elementType">The element type or IEnumerable<T> of the given type.</param>
        /// <returns>True if the type has an element type, else false.
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
