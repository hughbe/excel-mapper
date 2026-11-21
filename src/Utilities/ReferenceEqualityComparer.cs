using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

/// <summary>
/// Compares objects by reference, ignoring any <see cref="object.Equals(object)"/> override. The
/// netstandard builds don't have <c>System.Collections.Generic.ReferenceEqualityComparer</c>.
/// </summary>
internal sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class
{
    public static ReferenceEqualityComparer<T> Instance { get; } = new();

    public bool Equals(T? x, T? y) => ReferenceEquals(x, y);

    public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
}
