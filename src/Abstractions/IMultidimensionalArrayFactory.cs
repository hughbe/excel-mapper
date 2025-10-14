namespace ExcelMapper.Abstractions;

/// <summary>
/// Factory for creating multidimensional arrays with a defined lifecycle.
/// </summary>
/// <typeparam name="T">The type of elements in the array.</typeparam>
/// <remarks>
/// <para>
/// This interface defines a lifecycle pattern for building multidimensional arrays:
/// <list type="number">
/// <item><description>Call <see cref="Begin"/> to start a new array and specify the dimensions.</description></item>
/// <item><description>Call <see cref="Set"/> to place items at specific multi-dimensional indices.</description></item>
/// <item><description>Call <see cref="End"/> to finalize and retrieve the completed array.</description></item>
/// <item><description>Optionally call <see cref="Reset"/> to clean up state (or begin a new array cycle).</description></item>
/// </list>
/// </para>
/// <para>
/// The factory is reusable: after calling <see cref="End"/>, you can call <see cref="Begin"/> again to create a new array.
/// </para>
/// <para>
/// Calling <see cref="Set"/> before <see cref="Begin"/>, or calling <see cref="Begin"/> without first calling <see cref="End"/>,
/// will throw an <see cref="ExcelMappingException"/>.
/// </para>
/// </remarks>
public interface IMultidimensionalArrayFactory<T>
{
    /// <summary>
    /// Begins a new multidimensional array creation operation with the specified dimensions.
    /// </summary>
    /// <param name="lengths">An array specifying the size of each dimension. Must not be null or empty, and all values must be non-negative.</param>
    /// <exception cref="System.ArgumentNullException">Thrown when <paramref name="lengths"/> is null.</exception>
    /// <exception cref="System.ArgumentException">Thrown when <paramref name="lengths"/> is empty.</exception>
    /// <exception cref="System.ArgumentOutOfRangeException">Thrown when any value in <paramref name="lengths"/> is negative.</exception>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="Begin"/> is called before <see cref="End"/> was called on a previous operation.</exception>
    /// <remarks>
    /// For example, passing [3, 4] creates a 3×4 two-dimensional array, while [2, 3, 4] creates a 2×3×4 three-dimensional array.
    /// </remarks>
    void Begin(int[] lengths);
    
    /// <summary>
    /// Sets an item at a specific multi-dimensional index in the array being built.
    /// </summary>
    /// <param name="indices">An array specifying the index in each dimension. Must not be null or empty, and all values must be non-negative and within bounds.</param>
    /// <param name="item">The item to place at the specified indices.</param>
    /// <exception cref="System.ArgumentNullException">Thrown when <paramref name="indices"/> is null.</exception>
    /// <exception cref="System.ArgumentException">Thrown when <paramref name="indices"/> length does not match the number of dimensions specified in <see cref="Begin"/>.</exception>
    /// <exception cref="System.ArgumentOutOfRangeException">Thrown when any index value is negative or exceeds the corresponding dimension's length.</exception>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="Set"/> is called before <see cref="Begin"/>.</exception>
    /// <remarks>
    /// For a 3×4 array, valid indices would be [0, 0] through [2, 3].
    /// </remarks>
    void Set(int[] indices, T? item);
    
    /// <summary>
    /// Completes the array creation and returns the finalized multidimensional array.
    /// </summary>
    /// <returns>The completed multidimensional array as an object. Cast to the appropriate array type (e.g., T[,] for 2D).</returns>
    /// <exception cref="ExcelMappingException">Thrown when <see cref="End"/> is called before <see cref="Begin"/>.</exception>
    /// <remarks>
    /// After calling this method, the factory is ready to begin a new array via <see cref="Begin"/>.
    /// </remarks>
    object End();
    
    /// <summary>
    /// Resets the factory to its initial state, clearing any in-progress array data.
    /// </summary>
    /// <remarks>
    /// This method can be called at any time and is safe to call multiple times.
    /// It is automatically called by <see cref="End"/>, so explicit calls are typically not necessary.
    /// </remarks>
    void Reset();
}
