using System.Collections.Concurrent;
using System.Linq;

namespace ExcelMapper.Binding;

/// <summary>
/// Setters and constructors an <see cref="ExcelMappingContext"/> registered (see
/// <see cref="ExcelTypeRegistrar.Setter{TOwner, TValue}"/> and
/// <see cref="ExcelTypeRegistrar.Constructor{TOwner}"/>), written by the source generator as
/// ordinary code. <see cref="CompiledAccessors"/> and <see cref="ConstructionPlan{T}"/> use them
/// rather than compiling expressions, which native AOT can't, or using reflection, which boxes.
/// </summary>
internal static class RegisteredAccessors
{
    private static readonly ConcurrentDictionary<(Type Owner, string Member, Type Value), Delegate> s_setters = new();
    // Constructors by their type and the assembly-qualified names of their parameter types.
    private static readonly ConcurrentDictionary<(Type Owner, string Parameters), Delegate> s_constructors = new();

    public static void RegisterSetter<TOwner, TValue>(string member, ExcelMemberSetter<TOwner, TValue> setter)
        => s_setters[(typeof(TOwner), member, typeof(TValue))] = setter;

    /// <summary>
    /// Gets the registered setter of a member of a <typeparamref name="TOwner"/>, whose type is
    /// <typeparamref name="TValue"/>.
    /// </summary>
    public static ExcelMemberSetter<TOwner, TValue>? GetSetter<TOwner, TValue>(string member)
        => s_setters.TryGetValue((typeof(TOwner), member, typeof(TValue)), out var setter) ? (ExcelMemberSetter<TOwner, TValue>)setter : null;

    public static void RegisterConstructor<TOwner>(Type[] parameterTypes, ExcelConstructor<TOwner> constructor)
        => s_constructors[(typeof(TOwner), Signature(parameterTypes))] = constructor;

    /// <summary>
    /// Gets the registered constructor of a <typeparamref name="TOwner"/> with the given parameter types.
    /// </summary>
    public static ExcelConstructor<TOwner>? GetConstructor<TOwner>(Type[] parameterTypes)
        => s_constructors.TryGetValue((typeof(TOwner), Signature(parameterTypes)), out var constructor) ? (ExcelConstructor<TOwner>)constructor : null;

    private static string Signature(Type[] parameterTypes) => string.Join(",", parameterTypes.Select(t => t.AssemblyQualifiedName));
}
