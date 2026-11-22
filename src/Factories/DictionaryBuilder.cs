using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// How a dictionary of one type is built from its keys and values, chosen once for a member by
/// <see cref="DictionaryRule{TKey, TValue}"/>. Like <see cref="CollectionBuilder{T}"/>, it holds no
/// state, so one builder is shared by every sheet and thread without a lock.
/// </summary>
internal sealed class DictionaryBuilder<TKey, TValue> where TKey : notnull
{
    private readonly Func<int, object> _begin;
    private readonly Action<object, TKey, TValue?> _add;
    private readonly Func<object, object>? _end;

    private DictionaryBuilder(Func<int, object> begin, Action<object, TKey, TValue?> add, Func<object, object>? end)
    {
        _begin = begin;
        _add = add;
        _end = end;
    }

    /// <summary>
    /// Starts a dictionary of <paramref name="count"/> values.
    /// </summary>
    /// <returns>What the values are added to.</returns>
    public object Begin(int count) => _begin(count);

    /// <summary>
    /// Adds a value.
    /// </summary>
    public void Add(object items, TKey key, TValue? value) => _add(items, key, value);

    /// <summary>
    /// Finishes the dictionary.
    /// </summary>
    /// <returns>The dictionary.</returns>
    public object End(object items) => _end is null ? items : _end(items);

    /// <summary>
    /// A <see cref="Dictionary{TKey, TValue}"/>.
    /// </summary>
    public static DictionaryBuilder<TKey, TValue> ForDictionary()
        => new(static count => new Dictionary<TKey, TValue?>(count), s_addToDictionary, null);

    /// <summary>
    /// A dictionary of a type with a parameterless constructor, created by
    /// <paramref name="create"/>, that values are added to with <paramref name="add"/>.
    /// </summary>
    public static DictionaryBuilder<TKey, TValue> Direct(Func<object> create, Action<object, TKey, TValue?> add)
        => new(_ => create(), add, null);

    /// <summary>
    /// A dictionary of a type with a parameterless constructor that implements
    /// <see cref="IDictionary{TKey, TValue}"/> or <see cref="IDictionary"/>.
    /// </summary>
    public static DictionaryBuilder<TKey, TValue>? ForInterfaces([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type dictionaryType, Func<object> create)
    {
        if (dictionaryType.ImplementsInterface(typeof(IDictionary<TKey, TValue>)))
        {
            return Direct(create, static (items, key, value) => ((IDictionary<TKey, TValue?>)items).Add(key, value));
        }

        if (dictionaryType.ImplementsInterface(typeof(IDictionary)))
        {
            return Direct(create, static (items, key, value) => ((IDictionary)items).Add(key, value));
        }

        return null;
    }

    /// <summary>
    /// A dictionary created from a <see cref="Dictionary{TKey, TValue}"/> of its values by
    /// <paramref name="create"/>, such as an immutable dictionary or one whose constructor takes the
    /// values. If <paramref name="copiesValues"/>, the dictionary doesn't keep the one it's created
    /// from, which is reused for the next dictionary built on the thread.
    /// </summary>
    public static DictionaryBuilder<TKey, TValue> FromDictionary(Func<Dictionary<TKey, TValue?>, object> create, bool copiesValues)
        => new(
            copiesValues ? static count => ValueBuffer.Rent(count) : static count => new Dictionary<TKey, TValue?>(count),
            s_addToDictionary,
            copiesValues
                ? items =>
                {
                    var dictionary = (Dictionary<TKey, TValue?>)items;
                    var result = create(dictionary);
                    ValueBuffer.Return(dictionary);
                    return result;
                }
                : items => create((Dictionary<TKey, TValue?>)items));

    private static readonly Action<object, TKey, TValue?> s_addToDictionary = static (items, key, value) => ((Dictionary<TKey, TValue?>)items).Add(key, value);

    /// <summary>
    /// A dictionary for each thread that dictionaries created from a copy of their values are
    /// gathered in, so that building one doesn't allocate a dictionary for every row.
    /// </summary>
    private static class ValueBuffer
    {
        // Larger dictionaries aren't kept, so that a thread doesn't hold on to a large row's values.
        private const int MaxKeptCount = 1024;

        [ThreadStatic]
        private static Dictionary<TKey, TValue?>? t_dictionary;

        public static Dictionary<TKey, TValue?> Rent(int count)
        {
            var dictionary = t_dictionary;
            t_dictionary = null;
            return dictionary ?? new Dictionary<TKey, TValue?>(count);
        }

        public static void Return(Dictionary<TKey, TValue?> dictionary)
        {
            if (dictionary.Count <= MaxKeptCount)
            {
                dictionary.Clear();
                t_dictionary = dictionary;
            }
        }
    }
}
