using System.ComponentModel;

namespace ExcelMapper;

/// <summary>
/// The arguments of a constructor called for a row (see <see cref="ExcelConstructor{TOwner}"/>):
/// the values the members that match its parameters read, or the parameters' default values.
/// </summary>
/// <typeparam name="TOwner">The type created.</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public readonly struct ExcelConstructorArguments<TOwner>
{
    private readonly BoundMember<TOwner>[] _members;
    private readonly BoundConstructorArgument[] _arguments;

    internal ExcelConstructorArguments(BoundMember<TOwner>[] members, BoundConstructorArgument[] arguments)
    {
        _members = members;
        _arguments = arguments;
    }

    /// <summary>
    /// Gets the argument of a parameter, as a <typeparamref name="T"/>: the value its member read,
    /// without boxing if the member's map reads a <typeparamref name="T"/>, or else the parameter's
    /// default value. A null value of a value type is its default value.
    /// </summary>
    /// <typeparam name="T">The parameter's type.</typeparam>
    /// <param name="index">The zero-based position of the parameter.</param>
    /// <returns>The argument.</returns>
    public T Get<T>(int index)
    {
        var argument = _arguments[index];
        if (argument.MemberIndex >= 0 && _members[argument.MemberIndex] is { HasValue: true } member)
        {
            return member is BoundMember<TOwner, T> typedMember
                ? BoundMember<TOwner, T>.GetValueOrDefault(typedMember, default!)
                : Convert<T>(member.Value);
        }

        return Convert<T>(argument.DefaultValue);
    }

    // As a compiled constructor converts an argument read as an object.
    private static T Convert<T>(object? value) => value is null ? default! : (T)value;
}
