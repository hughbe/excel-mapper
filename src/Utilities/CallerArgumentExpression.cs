#if NETSTANDARD
namespace System.Runtime.CompilerServices;

/// <summary>
/// Polyfill so the compiler fills in parameter names for <c>ThrowHelpers</c> on netstandard,
/// which doesn't include this attribute. The compiler only recognizes it in this namespace.
/// </summary>
[AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
internal sealed class CallerArgumentExpressionAttribute : Attribute
{
    public CallerArgumentExpressionAttribute(string parameterName)
    {
        ParameterName = parameterName;
    }

    public string ParameterName { get; }
}
#endif
