namespace ExcelMapper.Abstractions;

/// <summary>
/// Factory for creating enumerable collections with a defined lifecycle.
/// </summary>
/// <typeparam name="T">The type of elements in the collection.</typeparam>
/// <remarks>
/// <para>
/// This interface defines a lifecycle pattern for building collections:
/// <list type="number">
/// <item><description>Call <see cref="Begin"/> to start a new collection and specify the expected item count.</description></item>
/// <item><description>Call <see cref="Add"/> to append items sequentially, or <see cref="Set"/> to place items at specific indices.</description></item>
/// <item><description>Call <see cref="End"/> to finalize and retrieve the completed collection.</description></item>
/// <item><description>Optionally call <see cref="Reset"/> to clean up state (or begin a new collection cycle).</description></item>
/// </list>
/// </para>
/// <para>
/// The factory is reusable: after calling <see cref="End"/>, you can call <see cref="Begin"/> again to create a new collection.
/// </para>
/// <para>
/// Calling <see cref="Add"/> or <see cref="Set"/> before <see cref="Begin"/>, or calling <see cref="Begin"/> without first calling <see cref="End"/>,
/// will throw an <see cref="ExcelMappingException"/>.
/// </para>
/// </remarks>
public interface IEnumerableFactory<T>
{
    /// <summary>
    /// Begins a new collection creation operation with the expected number of items.
    /// </summary>
    /// <param name="count">The expected number of items that will be added to the collection. Must be non-negative.</param>
    /// <exception cref="System.ArgumentOutOfRangeException">Thrown when <paramref name="count"/> is negative.</exception>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="Begin"/> is called before <see cref="End"/> was called on a previous operation.</exception>
    void Begin(int count);
    
    /// <summary>
    /// Adds an item to the collection being built.
    /// </summary>
    /// <param name="item">The item to add to the collection.</param>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="Add"/> is called before <see cref="Begin"/>.</exception>
    /// <remarks>
    /// Items are typically added sequentially. For some implementations (like arrays), adding beyond the initial count may throw an exception.
    /// </remarks>
    void Add(T? item);
    
    /// <summary>
    /// Sets an item at a specific index in the collection being built.
    /// </summary>
    /// <param name="index">The zero-based index at which to place the item. Must be non-negative.</param>
    /// <param name="item">The item to place at the specified index.</param>
    /// <exception cref="System.ArgumentOutOfRangeException">Thrown when <paramref name="index"/> is negative.</exception>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="Set"/> is called before <see cref="Begin"/>.</exception>
    /// <exception cref="System.NotSupportedException">Thrown by implementations that do not support indexed access (e.g., sets).</exception>
    /// <remarks>
    /// For list-based implementations, if the index exceeds the current count, the collection may be automatically grown.
    /// For fixed-size implementations like arrays, setting an index beyond the initial size may throw an exception.
    /// </remarks>
    void Set(int index, T? item);
    
    /// <summary>
    /// Completes the collection creation and returns the finalized collection.
    /// </summary>
    /// <returns>The completed collection as an object. The actual type depends on the implementation.</returns>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="End"/> is called before <see cref="Begin"/>.</exception>
    /// <remarks>
    /// After calling this method, the factory is ready to begin a new collection via <see cref="Begin"/>.
    /// The returned object should be cast to the appropriate collection type by the caller.
    /// </remarks>
    object End();
    
    /// <summary>
    /// Resets the factory to its initial state, clearing any in-progress collection data.
    /// </summary>
    /// <remarks>
    /// This method can be called at any time and is safe to call multiple times.
    /// It is automatically called by <see cref="End"/>, so explicit calls are typically not necessary.
    /// </remarks>
    void Reset();
}
