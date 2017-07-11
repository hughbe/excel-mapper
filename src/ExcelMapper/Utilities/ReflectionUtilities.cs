using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Utilities
{
    internal static class ReflectionUtilities
    {
        public static Type GetIEnumerableType(this Type type)
        {
            if (type.IsGenericInterface(typeof(IEnumerable<>)))
            {
                return type.GenericTypeArguments[0];
            }

            return type.GetTypeInfo().ImplementedInterfaces.GetIEnumerableElementType();            
        }

        public static Type GetIEnumerableElementType(this IEnumerable<Type> types)
        {
            Type ienumerableType = types.FirstOrDefault(parameterType => parameterType.IsGenericInterface(typeof(IEnumerable<>)));
            return ienumerableType?.GenericTypeArguments[0];
        }

        public static bool ImplementsInterface(this Type type, Type interfaceType)
        {
            return type.GetTypeInfo().ImplementedInterfaces.Any(t => t == interfaceType);
        }

        public static bool IsGenericInterface(this Type type, Type interfaceType)
        {
            if (!type.GetTypeInfo().IsGenericType)
            {
                return false;
            }

            return type.GetGenericTypeDefinition() == interfaceType;
        }

        public static bool IsAssignableFrom(this Type type, Type other)
        {
            return type.GetTypeInfo().IsAssignableFrom(other.GetTypeInfo());
        }

        public static ConstructorInfo GetConstructor(this Type type, Type[] parameters)
        {
            IEnumerable<ConstructorInfo> constructors = type.GetTypeInfo().DeclaredConstructors;
            return constructors.First(c => c.HasParameters(parameters));
        }

        private static bool HasParameters(this MethodBase method, Type[] parameters)
        {
            IEnumerable<Type> parameterTypes = method.GetParameters().Select(p => p.ParameterType);
            return parameterTypes.SequenceEqual(parameters);
        }

        public static object DefaultValue(this Type type)
        {
            if (type.GetTypeInfo().IsValueType)
            {
                return Activator.CreateInstance(type);
            }

            return null;
        }
    }
}
