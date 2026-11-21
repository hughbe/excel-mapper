using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a set by instantiating a type that implements ISet&lt;T&gt; and adding items.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
internal sealed class ISetTImplementingEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private readonly Func<object> _create;
    private ISet<T?>? _items;

    /// <summary>
    /// Constructs a factory that creates sets of the given type.
    /// </summary>
    /// <param name="setType">The type of set to create.</param>
    /// <exception cref="ArgumentException">Thrown when the set type is invalid or unsupported.</exception>
    public ISetTImplementingEnumerableFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type setType)
    {
        _create = CompiledAccessors.GetParameterlessConstructor(setType);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = (ISet<T?>)_create();
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _items!.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        throw new NotSupportedException($"Set is not supported for {nameof(HashSetEnumerableFactory<T>)}.");
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
        var factory = (ISetTImplementingEnumerableFactory<T>)MemberwiseClone();
        factory.ResetForBinding();
        return factory;
    }

    // A binding's copy starts with nothing being built and no buffer of its own. Only drop the
    // references: the objects may be in use by the map's own factory on another thread.
    private void ResetForBinding()
    {
        _items = null;
    }

    public void Reset()
    {
        _items = null;
    }
}
