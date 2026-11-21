#if !NET5_0_OR_GREATER
// Polyfills so the netstandard builds compile with the trimming annotations used by the net8.0
// build. They're internal: the trimmer and analyzers only run against the net8.0 build.
namespace System.Diagnostics.CodeAnalysis;

[Flags]
internal enum DynamicallyAccessedMemberTypes
{
    None = 0,
    PublicParameterlessConstructor = 0x0001,
    PublicConstructors = 0x0002 | PublicParameterlessConstructor,
    NonPublicConstructors = 0x0004,
    PublicMethods = 0x0008,
    NonPublicMethods = 0x0010,
    PublicFields = 0x0020,
    NonPublicFields = 0x0040,
    PublicNestedTypes = 0x0080,
    NonPublicNestedTypes = 0x0100,
    PublicProperties = 0x0200,
    NonPublicProperties = 0x0400,
    PublicEvents = 0x0800,
    NonPublicEvents = 0x1000,
    Interfaces = 0x2000,
    All = ~None
}

[AttributeUsage(
    AttributeTargets.Field | AttributeTargets.ReturnValue | AttributeTargets.GenericParameter |
    AttributeTargets.Parameter | AttributeTargets.Property | AttributeTargets.Method |
    AttributeTargets.Class | AttributeTargets.Interface | AttributeTargets.Struct,
    Inherited = false)]
internal sealed class DynamicallyAccessedMembersAttribute(DynamicallyAccessedMemberTypes memberTypes) : Attribute
{
    public DynamicallyAccessedMemberTypes MemberTypes { get; } = memberTypes;
}

[AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class, Inherited = false)]
internal sealed class RequiresUnreferencedCodeAttribute(string message) : Attribute
{
    public string Message { get; } = message;

    public string? Url { get; set; }
}

[AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class, Inherited = false)]
internal sealed class RequiresDynamicCodeAttribute(string message) : Attribute
{
    public string Message { get; } = message;

    public string? Url { get; set; }
}

[AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = true)]
internal sealed class UnconditionalSuppressMessageAttribute(string category, string checkId) : Attribute
{
    public string Category { get; } = category;

    public string CheckId { get; } = checkId;

    public string? Justification { get; set; }
}
#endif
