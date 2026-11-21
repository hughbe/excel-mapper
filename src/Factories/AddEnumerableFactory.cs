using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a collection by calling the Add method to add items.
/// </summary>
/// <typeparam name="T">The type of the collection items.</typeparam>
internal sealed class AddEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private object? _items;
    private readonly Func<object> _create;
    private readonly Action<object, T?> _add;

    /// <summary>
    /// Constructs a factory that creates collections of the given type.
    /// </summary>
    /// <param name="collectionType">The type of collection to create.</param>
    /// <exception cref="ArgumentException">Thrown when the collection type is invalid or unsupported.</exception>
    public AddEnumerableFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type collectionType)
    {
        var addMethod = collectionType.GetMethod("Add", [typeof(T)]) ?? throw new ArgumentException($"Type does not have an Add({typeof(T)}) method.", nameof(collectionType));
        _create = CompiledAccessors.GetParameterlessConstructor(collectionType);
        _add = CompiledAccessors.GetAction<T?>(addMethod);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = _create();
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _add(_items!, item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        throw new NotSupportedException($"Set is not supported for {nameof(AddEnumerableFactory<T>)}.");
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
        var factory = (AddEnumerableFactory<T>)MemberwiseClone();
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
