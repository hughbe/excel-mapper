using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Builders;

/// <summary>
/// How a collection of one type is built from its elements, chosen once for a member by
/// <see cref="CollectionRule{TElement}"/>. A builder holds no state: the collection being built is
/// passed to each call, so one builder is shared by every sheet and thread without a lock.
/// </summary>
/// <remarks>
/// Collections are either built directly, adding each element to the result (arrays, lists, sets
/// and types with an <c>Add</c> method), or gathered into a list that the result is created from
/// (immutable and frozen collections, and types created by a constructor that takes the elements).
/// </remarks>
internal sealed class CollectionBuilder<T>
{
    private readonly Func<int, object> _begin;
    private readonly Action<object, int, T?> _add;
    private readonly Action<object, int, T?>? _set;
    private readonly Func<object, object>? _end;

    private CollectionBuilder(Func<int, object> begin, Action<object, int, T?> add, Action<object, int, T?>? set, Func<object, object>? end)
    {
        _begin = begin;
        _add = add;
        _set = set;
        _end = end;
    }

    /// <summary>
    /// Starts a collection of <paramref name="count"/> elements.
    /// </summary>
    /// <returns>What the elements are added to.</returns>
    public object Begin(int count) => _begin(count);

    /// <summary>
    /// Adds the next element, the <paramref name="index"/>th.
    /// </summary>
    public void Add(object items, int index, T? item) => _add(items, index, item);

    /// <summary>
    /// Sets the element at <paramref name="index"/>, for elements mapped one by one, such as
    /// <c>Map(p => p.Values[0])</c>. Lists grow to fit.
    /// </summary>
    public void Set(object items, int index, T? item)
        => (_set ?? throw new NotSupportedException("Elements of this collection can't be set by index."))(items, index, item);

    /// <summary>
    /// Finishes the collection.
    /// </summary>
    /// <returns>The collection.</returns>
    public object End(object items) => _end is null ? items : _end(items);

    /// <summary>
    /// An array, created at its final size.
    /// </summary>
    public static CollectionBuilder<T> ForArray()
        => new(
            static count => new T?[count],
            static (items, index, item) => ((T?[])items)[index] = item,
            static (items, index, item) => ((T?[])items)[index] = item,
            null);

    /// <summary>
    /// A <see cref="List{T}"/>.
    /// </summary>
    public static CollectionBuilder<T> ForList()
        => new(static count => new List<T?>(count), s_addToList, s_setInList, null);

    /// <summary>
    /// A <see cref="HashSet{T}"/>.
    /// </summary>
    public static CollectionBuilder<T> ForHashSet()
        => new(static count => CreateHashSet(count), static (items, _, item) => ((HashSet<T?>)items).Add(item), null, null);

    /// <summary>
    /// A collection created by <paramref name="create"/> that elements are added to with
    /// <paramref name="add"/>, and set by index with <paramref name="set"/> if it has an indexer.
    /// </summary>
    public static CollectionBuilder<T> Direct(Func<object> create, Action<object, T?> add, Action<object, int, T?>? set = null)
        => new(_ => create(), (items, _, item) => add(items, item), set, null);

    /// <summary>
    /// A collection of a type with a parameterless constructor that implements <see cref="IList{T}"/>,
    /// <see cref="ISet{T}"/>, <see cref="ICollection{T}"/> or <see cref="IList"/>.
    /// </summary>
    public static CollectionBuilder<T>? ForInterfaces([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type collectionType, Func<object> create)
    {
        if (collectionType.ImplementsInterface(typeof(IList<T>)))
        {
            return Direct(create, static (items, item) => ((IList<T?>)items).Add(item), static (items, index, item) => SetInList((IList<T?>)items, index, item));
        }

        if (collectionType.ImplementsInterface(typeof(ISet<T>)))
        {
            return Direct(create, static (items, item) => ((ISet<T?>)items).Add(item));
        }

        if (collectionType.ImplementsInterface(typeof(ICollection<T>)))
        {
            return Direct(create, static (items, item) => ((ICollection<T?>)items).Add(item));
        }

        if (collectionType.ImplementsInterface(typeof(IList)))
        {
            return Direct(create, static (items, item) => ((IList)items).Add(item), static (items, index, item) =>
            {
                var list = (IList)items;
                while (list.Count <= index)
                {
                    list.Add(default(T));
                }

                list[index] = item;
            });
        }

        return null;
    }

    /// <summary>
    /// A collection created from a list of its elements by <paramref name="create"/>, such as an
    /// immutable collection or one whose constructor takes the elements. If
    /// <paramref name="copiesElements"/>, the collection doesn't keep the list, so the list is reused
    /// for the next collection built on the thread.
    /// </summary>
    public static CollectionBuilder<T> FromList(Func<List<T?>, object> create, bool copiesElements)
        => new(
            copiesElements ? static count => ElementBuffer.Rent(count) : static count => new List<T?>(count),
            s_addToList,
            s_setInList,
            copiesElements
                ? items =>
                {
                    var list = (List<T?>)items;
                    var result = create(list);
                    ElementBuffer.Return(list);
                    return result;
                }
                : items => create((List<T?>)items));

    /// <summary>
    /// A collection created from a set of its elements by <paramref name="create"/>, such as
    /// <c>ReadOnlySet&lt;T&gt;(ISet&lt;T&gt;)</c>.
    /// </summary>
    public static CollectionBuilder<T> FromSet(Func<object, object> create)
        => new(static count => CreateHashSet(count), static (items, _, item) => ((HashSet<T?>)items).Add(item), null, create);

    private static readonly Action<object, int, T?> s_addToList = static (items, _, item) => ((List<T?>)items).Add(item);

    private static readonly Action<object, int, T?> s_setInList = static (items, index, item) => SetInList((List<T?>)items, index, item);

    // Lists mapped by index grow to fit.
    private static void SetInList(IList<T?> list, int index, T? item)
    {
        while (list.Count <= index)
        {
            list.Add(default);
        }

        list[index] = item;
    }

    private static HashSet<T?> CreateHashSet(int count)
    {
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP2_1_OR_GREATER
        return new HashSet<T?>(count);
#else
        return [];
#endif
    }

    /// <summary>
    /// A list for each thread that collections created from a copy of their elements are gathered
    /// in, so that building one doesn't allocate a list for every row.
    /// </summary>
    private static class ElementBuffer
    {
        // Larger lists aren't kept, so that a thread doesn't hold on to a large row's elements.
        private const int MaxKeptCapacity = 1024;

        [ThreadStatic]
        private static List<T?>? t_list;

        public static List<T?> Rent(int count)
        {
            // Take the list, so that a collection built while this one is (by a converter, say)
            // gets its own.
            var list = t_list;
            t_list = null;
            return list ?? new List<T?>(count);
        }

        public static void Return(List<T?> list)
        {
            if (list.Capacity <= MaxKeptCapacity)
            {
                list.Clear();
                t_list = list;
            }
        }
    }
}
