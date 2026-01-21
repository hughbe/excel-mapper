using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a multidimensional array by pre-allocating storage with specified dimensions and setting values by indices.
/// </summary>
/// <typeparam name="T">The type of the array elements.</typeparam>
public class MultidimensionalArrayFactory<T> : IMultidimensionalArrayFactory<T>
{
    private Array? _items;

    /// <inheritdoc/>
    public void Begin(int[] lengths)
    {
        ThrowHelpers.ThrowIfNull(lengths, nameof(lengths));
        if (lengths.Length == 0)
        {
            throw new ArgumentException("Lengths cannot be empty.", nameof(lengths));
        }
        foreach (var length in lengths)
        {
            ThrowHelpers.ThrowIfNegative(length, nameof(lengths));
        }

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = Array.CreateInstance(typeof(T), lengths);
    }

    /// <inheritdoc/>
    public void Set(int[] indices, T? item)
    {
        ThrowHelpers.ThrowIfNull(indices, nameof(indices));
        if (indices.Length == 0)
        {
            throw new ArgumentException("Indices cannot be empty.", nameof(indices));
        }
        foreach (var index in indices)
        {
            ThrowHelpers.ThrowIfNegative(index, nameof(indices));
        }

        EnsureMapping();
        _items.SetValue(item, indices);
    }

    /// <inheritdoc/>
    public object End()
    {
        EnsureMapping();

        try
        {
            return _items;
        }
        finally
        {
            Reset();
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        _items = null;
    }

    [MemberNotNull(nameof(_items))]
    private void EnsureMapping()
    {
        if (_items is null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
