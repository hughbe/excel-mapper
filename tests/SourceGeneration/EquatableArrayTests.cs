using System.Collections;
using System.Collections.Immutable;
using ExcelMapper.SourceGeneration;

namespace ExcelMapper.Tests;

/// <summary>
/// The generator caches its models between compilations, comparing them by value: an array of the
/// same items is equal, with the same hash code.
/// </summary>
public class EquatableArrayTests
{
    [Fact]
    public void Equals_SameItems_ReturnsTrue()
    {
        var array = new EquatableArray<string>(["a", "b"]);
        var other = new EquatableArray<string>(["a", "b"]);
        Assert.True(array.Equals(other));
        Assert.True(array.Equals((object)other));
        Assert.Equal(array.GetHashCode(), other.GetHashCode());
    }

    [Fact]
    public void Equals_DifferentItems_ReturnsFalse()
    {
        var array = new EquatableArray<string>(["a", "b"]);
        Assert.False(array.Equals(new EquatableArray<string>(["a"])));
        Assert.False(array.Equals(new EquatableArray<string>(["b", "a"])));
        Assert.False(array.Equals((object)"a"));
    }

    [Fact]
    public void Default_IsEmpty()
    {
        var array = default(EquatableArray<string>);
        Assert.Empty(array);
        Assert.True(array.Equals(new EquatableArray<string>(ImmutableArray<string>.Empty)));
        Assert.False(((IEnumerable)array).GetEnumerator().MoveNext());
    }
}
