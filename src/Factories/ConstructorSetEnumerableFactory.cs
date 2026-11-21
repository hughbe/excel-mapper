using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a set by passing a HashSet&lt;T&gt; of items to a constructor that accepts ISet&lt;T&gt;.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
internal sealed class ConstructorSetEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private readonly Func<object, object> _construct;

    private HashSet<T?>? _items;

    /// <summary>
    /// Constructs a factory that creates sets of the given type.
    /// </summary>
    /// <param name="setType">The type of set to create.</param>
    /// <exception cref="ArgumentException">Thrown when the set type is invalid or unsupported.</exception>
    public ConstructorSetEnumerableFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type setType)
    {
        var constructor = CollectionUtilities.GetSetConstructor<T>(setType)!;

        _construct = CompiledAccessors.GetConstructor1(constructor);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP2_1_OR_GREATER
        _items = new HashSet<T?>(count);
#else
        _items = new HashSet<T?>();
#endif
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _items!.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        throw new NotSupportedException($"Set is not supported for {nameof(ConstructorSetEnumerableFactory<T>)}.");
    }

    /// <inheritdoc/>
    public object End()
    {
        try
        {
            return _construct(_items!);
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
        var factory = (ConstructorSetEnumerableFactory<T>)MemberwiseClone();
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
