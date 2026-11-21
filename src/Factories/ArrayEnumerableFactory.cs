using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an array by pre-allocating storage and adding items sequentially.
/// </summary>
/// <typeparam name="T">The type of the array elements.</typeparam>
internal sealed class ArrayEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private int _currentIndex = -1;
    private T?[]? _items;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = new T[count];
        _currentIndex = 0;
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _items![_currentIndex++] = item;
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        ThrowHelpers.ThrowIfGreaterThanOrEqual(index, _items!.Length, nameof(index));

        _items[index] = item;
    }

    /// <inheritdoc/>
    public object End()
    {
        try
        {
            return _items!;
        }
        finally
        {
            Reset();
        }
    }

    /// <inheritdoc/>
    /// <inheritdoc/>
    object IBindingFactory.CreateForBinding()
    {
        var factory = (ArrayEnumerableFactory<T>)MemberwiseClone();
        factory.ResetForBinding();
        return factory;
    }

    // A binding's copy starts with nothing being built and no buffer of its own. Only drop the
    // references: the objects may be in use by the map's own factory on another thread.
    private void ResetForBinding()
    {
        _currentIndex = -1;
    }

    public void Reset()
    {
        _items = null;
        _currentIndex = -1;
    }
}
