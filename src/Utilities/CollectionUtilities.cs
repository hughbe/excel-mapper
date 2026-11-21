using System.Collections.Concurrent;

namespace ExcelMapper.Utilities;

internal static class CollectionUtilities
{
    /// <summary>
    /// Whether a collection type's constructor copies the collection passed to it, rather than keeping
    /// it (as <see cref="System.Collections.ObjectModel.ReadOnlyCollection{T}"/> does). Only these
    /// types can be given a buffer that is reused for the next row. Types not listed (including the
    /// user's own types) get a new collection for each row.
    /// </summary>
    public static bool ConstructorCopies(Type type)
    {
        if (!type.IsGenericType)
        {
            return false;
        }

        var definition = type.GetGenericTypeDefinition();
        return definition == typeof(Queue<>)
            || definition == typeof(Stack<>)
            || definition == typeof(LinkedList<>)
            || definition == typeof(SortedSet<>)
            || definition == typeof(HashSet<>)
            || definition == typeof(List<>)
            || definition == typeof(ConcurrentQueue<>)
            || definition == typeof(ConcurrentStack<>)
            || definition == typeof(ConcurrentBag<>)
            || definition == typeof(System.Collections.ObjectModel.ObservableCollection<>)
            || definition == typeof(Dictionary<,>)
            || definition == typeof(SortedDictionary<,>)
            || definition == typeof(SortedList<,>);
    }
}
