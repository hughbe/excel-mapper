namespace ExcelMapper.Abstractions;

/// <summary>
/// Factory for creating dictionary collections with a defined lifecycle.
/// </summary>
/// <typeparam name="TKey">The type of keys in the dictionary.</typeparam>
/// <typeparam name="TValue">The type of values in the dictionary.</typeparam>
/// <remarks>
/// <para>
/// This interface defines a lifecycle pattern for building dictionaries:
/// <list type="number">
/// <item><description>Call <see cref="Begin"/> to start a new dictionary and specify the expected number of entries.</description></item>
/// <item><description>Call <see cref="Add"/> to insert key-value pairs into the dictionary.</description></item>
/// <item><description>Call <see cref="End"/> to finalize and retrieve the completed dictionary.</description></item>
/// <item><description>Optionally call <see cref="Reset"/> to clean up state (or begin a new dictionary cycle).</description></item>
/// </list>
/// </para>
/// <para>
/// The factory is reusable: after calling <see cref="End"/>, you can call <see cref="Begin"/> again to create a new dictionary.
/// </para>
/// <para>
/// Calling <see cref="Add"/> before <see cref="Begin"/>, or calling <see cref="Begin"/> without first calling <see cref="End"/>,
/// will throw an <see cref="ExcelMappingException"/>.
/// </para>
/// </remarks>
public interface IDictionaryFactory<TKey, TValue>
{
    /// <summary>
    /// Begins a new dictionary creation operation with the expected number of entries.
    /// </summary>
    /// <param name="count">The expected number of key-value pairs that will be added to the dictionary. Must be non-negative.</param>
    /// <exception cref="System.ArgumentOutOfRangeException">Thrown when <paramref name="count"/> is negative.</exception>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="Begin"/> is called before <see cref="End"/> was called on a previous operation.</exception>
    void Begin(int count);
    
    /// <summary>
    /// Adds a key-value pair to the dictionary being built.
    /// </summary>
    /// <param name="key">The key of the entry to add. Must not be null.</param>
    /// <param name="value">The value associated with the key.</param>
    /// <exception cref="System.ArgumentNullException">Thrown when <paramref name="key"/> is null.</exception>
    /// <exception cref="System.ArgumentException">Thrown when a key with the same value already exists in the dictionary.</exception>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="Add"/> is called before <see cref="Begin"/>.</exception>
    void Add(TKey key, TValue? value);
    
    /// <summary>
    /// Completes the dictionary creation and returns the finalized dictionary.
    /// </summary>
    /// <returns>The completed dictionary as an object. The actual type depends on the implementation.</returns>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="End"/> is called before <see cref="Begin"/>.</exception>
    /// <remarks>
    /// After calling this method, the factory is ready to begin a new dictionary via <see cref="Begin"/>.
    /// The returned object should be cast to the appropriate dictionary type by the caller.
    /// </remarks>
    object End();
    
    /// <summary>
    /// Resets the factory to its initial state, clearing any in-progress dictionary data.
    /// </summary>
    /// <remarks>
    /// This method can be called at any time and is safe to call multiple times.
    /// It is automatically called by <see cref="End"/>, so explicit calls are typically not necessary.
    /// </remarks>
    void Reset();
}
