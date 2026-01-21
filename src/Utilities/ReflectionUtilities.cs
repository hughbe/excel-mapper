using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;

namespace ExcelMapper.Utilities;

internal static class ReflectionUtilities
{
#if !NET5_0_OR_GREATER
    public static bool IsAssignableTo(this Type type, Type targetType)
    {
        return targetType.GetTypeInfo().IsAssignableFrom(type.GetTypeInfo());
    }
#endif

    public static bool ImplementsInterface(this Type type, Type interfaceType)
    {
        return type.GetTypeInfo().ImplementedInterfaces.Any(t => t == interfaceType);
    }

    public static bool ImplementsGenericInterface(
        this Type type,
        Type genericInterfaceType,
        [NotNullWhen(true)] out Type? elementType)
    {
        bool CheckInterface(Type interfaceType, [NotNullWhen(true)] out Type? elementTypeResult)
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

    /// <summary>
    /// Gets the element type or the IEnumerable&lt;T&gt; type of the given type.
    /// </summary>
    /// <param name="type">The type to get the element type of.</param>
    /// <param name="elementType">The element type or IEnumerable&lt;T&gt; of the given type.</param>
    public static Type? GetElementTypeOrEnumerableType(this Type type)
    {
        // Array type.
        if (type.IsArray)
        {
            return type.GetElementType()!;
        }

        // Generic lists use type object.
        if (type.ImplementsGenericInterface(typeof(IEnumerable<>), out var elementType))
        {
            return elementType;
        }

        // Non-generic interfaces use type object.
        if (type == typeof(IEnumerable) || type.ImplementsInterface(typeof(IEnumerable)))
        {
            return typeof(object);
        }

        return null;
    }
    
    [ExcludeFromCodeCoverage]
    public static object InvokeUnwrapped(this MethodInfo method, object? obj, params object?[] parameters)
    {
        try
        {
            return method.Invoke(obj, parameters)!;
        }
        catch (TargetInvocationException ex)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException!).Throw();
            throw; // Will never be hit, but compiler doesn't know that.
        }
    }
}
