using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a multidimensional array by pre-allocating storage with specified dimensions and setting values by indices.
/// </summary>
/// <typeparam name="T">The type of the array elements.</typeparam>
internal sealed class MultidimensionalArrayFactory<T> : IMultidimensionalArrayFactory<T>
{
    private Array? _items;

    /// <summary>
    /// Initializes a new instance of <see cref="MultidimensionalArrayFactory{T}"/>.
    /// </summary>
    [RequiresDynamicCode("Creates multidimensional arrays of T whose rank is only known at run time, which native AOT can't guarantee.")]
    public MultidimensionalArrayFactory()
    {
    }

    /// <inheritdoc/>
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "The constructor requires dynamic code.")]
    public void Begin(int[] lengths)
    {
        foreach (var length in lengths)
        {
        }

        _items = Array.CreateInstance(typeof(T), lengths);
    }

    /// <inheritdoc/>
    public void Set(int[] indices, T? item)
    {
        foreach (var index in indices)
        {
        }

        _items!.SetValue(item, indices);
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
    public void Reset()
    {
        _items = null;
    }
}
