using System.Collections;
using System.Collections.Frozen;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Factories;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// The rule for a type with elements (an array, list, set or other collection): how its elements
/// are mapped, and how a collection of its type is created. See <see cref="TypeRule"/>.
/// </summary>
[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal abstract class CollectionRule
{
    /// <summary>
    /// Auto-maps a member of the collection type to a <see cref="CollectionMap{TElement}"/> that
    /// reads its elements from several cells.
    /// </summary>
    public abstract bool TryCreateMap(MemberInfo? member, Type collectionType, ICellsReaderFactory readerFactory, EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map);

    /// <summary>
    /// Gets the map for elements mapped one by one, such as <c>Map(p => p.Values[0])</c>.
    /// </summary>
    public abstract IMap GetOrCreateIndexerMap(IMap parentMap, Type collectionType, object? context);

    /// <summary>
    /// Gets the map for the elements of a multidimensional array mapped one by one, such as
    /// <c>Map(p => p.Values[0, 1])</c>.
    /// </summary>
    public abstract IMap GetOrCreateMultidimensionalIndexerMap(IMap parentMap, object? context);
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal sealed class CollectionRule<TElement> : CollectionRule
{
    public override bool TryCreateMap(MemberInfo? member, Type collectionType, ICellsReaderFactory readerFactory, EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        if (TryCreateCollectionMap(member, collectionType, readerFactory, emptyValueStrategy, out var collectionMap))
        {
            map = collectionMap;
            return true;
        }

        map = null;
        return false;
    }

    /// <summary>
    /// Creates a <see cref="CollectionMap{TElement}"/> for a member of type
    /// <paramref name="collectionType"/>, with the element type's built-in conversion. Fails if the
    /// collection can't be created or its elements have no conversion.
    /// </summary>
    public static bool TryCreateCollectionMap(MemberInfo? member, Type collectionType, ICellsReaderFactory readerFactory, EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out CollectionMap<TElement>? map)
    {
        if (!TryGetFactory(collectionType, out var factory))
        {
            map = null;
            return false;
        }

        map = new CollectionMap<TElement>(readerFactory, factory);
        if (!AutoMapper.TrySetupMap<TElement>(member, map, emptyValueStrategy, isAutoMapping: true))
        {
            map = null;
            return false;
        }

        return true;
    }

    public override IMap GetOrCreateIndexerMap(IMap parentMap, Type collectionType, object? context)
    {
        if (AutoMapper.GetExistingMap(parentMap, context) is { } existingMap)
        {
            return existingMap as CollectionIndexerMap<TElement>
                ?? throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
        }

        if (!TryGetFactory(collectionType, out var factory))
        {
            throw new ExcelMappingException($"Could not map array of type \"{collectionType}\".");
        }

        var map = new CollectionIndexerMap<TElement>(factory);
        AutoMapper.AddExistingMap(parentMap, context, map);
        return map;
    }

    public override IMap GetOrCreateMultidimensionalIndexerMap(IMap parentMap, object? context)
    {
        if (AutoMapper.GetExistingMap(parentMap, context) is { } existingMap)
        {
            return existingMap as MultidimensionalIndexerMap<TElement>
                ?? throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
        }

        var map = new MultidimensionalIndexerMap<TElement>(new MultidimensionalArrayFactory<TElement>());
        AutoMapper.AddExistingMap(parentMap, context, map);
        return map;
    }

    /// <summary>
    /// Finds how to create a collection of type <paramref name="collectionType"/> and add elements
    /// to it.
    /// </summary>
    private static bool TryGetFactory(Type collectionType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        if (TryGetWellKnownFactory(collectionType, out result))
        {
            return true;
        }

        if (collectionType.IsInterface)
        {
            return TryGetInterfaceFactory(collectionType, out result);
        }

        if (!collectionType.IsAbstract)
        {
            return TryGetConcreteTypeFactory(collectionType, out result);
        }

        result = default;
        return false;
    }

    private static bool TryGetWellKnownFactory(Type collectionType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        if (collectionType == typeof(Array) || (collectionType.IsArray && collectionType.GetArrayRank() == 1))
        {
            result = new ArrayEnumerableFactory<TElement>();
            return true;
        }
        if (collectionType == typeof(ImmutableArray<TElement>))
        {
            result = new ImmutableArrayEnumerableFactory<TElement>();
            return true;
        }
        if (collectionType == typeof(ImmutableList<TElement>) || collectionType == typeof(IImmutableList<TElement>))
        {
            result = new ImmutableListEnumerableFactory<TElement>();
            return true;
        }
        if (collectionType == typeof(ImmutableStack<TElement>) || collectionType == typeof(IImmutableStack<TElement>))
        {
            result = new ImmutableStackEnumerableFactory<TElement>();
            return true;
        }
        if (collectionType == typeof(ImmutableQueue<TElement>) || collectionType == typeof(IImmutableQueue<TElement>))
        {
            result = new ImmutableQueueEnumerableFactory<TElement>();
            return true;
        }
        if (collectionType == typeof(ImmutableSortedSet<TElement>))
        {
            result = new ImmutableSortedSetEnumerableFactory<TElement>();
            return true;
        }
        if (collectionType == typeof(ImmutableHashSet<TElement>) || collectionType == typeof(IImmutableSet<TElement>))
        {
            result = new ImmutableHashSetEnumerableFactory<TElement>();
            return true;
        }
        if (collectionType == typeof(FrozenSet<TElement>))
        {
            result = new FrozenSetEnumerableFactory<TElement>();
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetInterfaceFactory(Type collectionType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        // Add values by creating a list and assigning to the property.
        if (typeof(List<TElement>).IsAssignableTo(collectionType))
        {
            result = new ListEnumerableFactory<TElement>();
            return true;
        }
        if (typeof(HashSet<TElement>).IsAssignableTo(collectionType))
        {
            result = new HashSetEnumerableFactory<TElement>();
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetConcreteTypeFactory(Type collectionType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        var hasDefaultConstructor = collectionType.GetConstructor([]) is not null;

        // A parameterless constructor and a collection interface.
        if (hasDefaultConstructor && TryGetDefaultConstructorFactory(collectionType, out result))
        {
            return true;
        }

        // A constructor that takes the elements.
        if (TryGetConstructorBasedFactory(collectionType, out result))
        {
            return true;
        }

        // A parameterless constructor and an Add method.
        if (hasDefaultConstructor && collectionType.GetMethod("Add", [typeof(TElement)]) is not null)
        {
            result = new AddEnumerableFactory<TElement>(collectionType);
            return true;
        }

        // A constructor that takes an ObservableCollection, such as ReadOnlyObservableCollection<T>.
        if (CollectionUtilities.GetConstructorTaking(collectionType, typeof(ObservableCollection<TElement>)) is not null)
        {
            result = new ReadOnlyObservableCollectionEnumerableFactory<TElement>(collectionType);
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetDefaultConstructorFactory(Type collectionType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        // Add values through IList<TElement>.Add(TElement item).
        if (collectionType.ImplementsInterface(typeof(IList<TElement>)))
        {
            result = new IListTImplementingEnumerableFactory<TElement>(collectionType);
            return true;
        }

        // Add values through ISet<TElement>.Add(TElement item).
        if (collectionType.ImplementsInterface(typeof(ISet<TElement>)))
        {
            result = new ISetTImplementingEnumerableFactory<TElement>(collectionType);
            return true;
        }

        // Add values through ICollection<TElement>.Add(TElement item).
        if (collectionType.ImplementsInterface(typeof(ICollection<TElement>)))
        {
            result = new ICollectionTImplementingEnumerableFactory<TElement>(collectionType);
            return true;
        }

        // Add values through IList.Add(object item).
        if (collectionType.ImplementsInterface(typeof(IList)))
        {
            result = new IListImplementingEnumerableFactory<TElement>(collectionType);
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetConstructorBasedFactory(Type collectionType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        // A constructor that needs a set, such as ReadOnlySet<T>(ISet<T>).
        if (CollectionUtilities.GetSetConstructor<TElement>(collectionType) is not null)
        {
            result = new ConstructorSetEnumerableFactory<TElement>(collectionType);
            return true;
        }

        // A constructor that takes the items as a list, such as ReadOnlyCollection<T>(IList<T>),
        // Queue<T>(IEnumerable<T>) or Stack<T>(IEnumerable<T>).
        if (CollectionUtilities.GetListConstructor<TElement>(collectionType) is not null)
        {
            result = new ConstructorEnumerableFactory<TElement>(collectionType);
            return true;
        }

        result = default;
        return false;
    }
}
