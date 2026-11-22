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
    public abstract bool TryCreateMap(MemberInfo? member, Type collectionType, Columns columns, EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map);

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
    public override bool TryCreateMap(MemberInfo? member, Type collectionType, Columns columns, EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        if (TryCreateCollectionMap(member, collectionType, columns, emptyValueStrategy, out var collectionMap))
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
    public static bool TryCreateCollectionMap(MemberInfo? member, Type collectionType, Columns columns, EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out CollectionMap<TElement>? map)
    {
        if (!TryGetBuilder(collectionType, out var builder))
        {
            map = null;
            return false;
        }

        map = new CollectionMap<TElement>(columns, builder);
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

        if (!TryGetBuilder(collectionType, out var builder))
        {
            throw new ExcelMappingException($"Could not map array of type \"{collectionType}\".");
        }

        var map = new CollectionIndexerMap<TElement>(builder);
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

        var map = new MultidimensionalIndexerMap<TElement>();
        AutoMapper.AddExistingMap(parentMap, context, map);
        return map;
    }

    /// <summary>
    /// Finds how to build a collection of type <paramref name="collectionType"/> from its elements.
    /// </summary>
    internal static bool TryGetBuilder(Type collectionType, [NotNullWhen(true)] out CollectionBuilder<TElement>? builder)
    {
        builder = GetWellKnownBuilder(collectionType)
            ?? (collectionType.IsInterface ? GetInterfaceBuilder(collectionType)
            : !collectionType.IsAbstract ? GetConcreteTypeBuilder(collectionType)
            : null);
        return builder is not null;
    }

    private static CollectionBuilder<TElement>? GetWellKnownBuilder(Type collectionType)
    {
        if (collectionType == typeof(Array) || (collectionType.IsArray && collectionType.GetArrayRank() == 1))
        {
            return CollectionBuilder<TElement>.ForArray();
        }
        if (collectionType == typeof(List<TElement>))
        {
            return CollectionBuilder<TElement>.ForList();
        }
        if (collectionType == typeof(HashSet<TElement>))
        {
            return CollectionBuilder<TElement>.ForHashSet();
        }

        // Immutable and frozen collections are created from a copy of the elements.
        if (collectionType == typeof(ImmutableArray<TElement>))
        {
            return CollectionBuilder<TElement>.FromList(static items => ImmutableArray.CreateRange(items), copiesElements: true);
        }
        if (collectionType == typeof(ImmutableList<TElement>) || collectionType == typeof(IImmutableList<TElement>))
        {
            return CollectionBuilder<TElement>.FromList(static items => ImmutableList.CreateRange(items), copiesElements: true);
        }
        if (collectionType == typeof(ImmutableStack<TElement>) || collectionType == typeof(IImmutableStack<TElement>))
        {
            return CollectionBuilder<TElement>.FromList(static items => ImmutableStack.CreateRange(items), copiesElements: true);
        }
        if (collectionType == typeof(ImmutableQueue<TElement>) || collectionType == typeof(IImmutableQueue<TElement>))
        {
            return CollectionBuilder<TElement>.FromList(static items => ImmutableQueue.CreateRange(items), copiesElements: true);
        }
        if (collectionType == typeof(ImmutableSortedSet<TElement>))
        {
            return CollectionBuilder<TElement>.FromList(static items => ImmutableSortedSet.CreateRange(items), copiesElements: true);
        }
        if (collectionType == typeof(ImmutableHashSet<TElement>) || collectionType == typeof(IImmutableSet<TElement>))
        {
            return CollectionBuilder<TElement>.FromList(static items => ImmutableHashSet.CreateRange(items), copiesElements: true);
        }
        if (collectionType == typeof(FrozenSet<TElement>))
        {
            return CollectionBuilder<TElement>.FromList(static items => items.ToFrozenSet(), copiesElements: true);
        }

        return null;
    }

    private static CollectionBuilder<TElement>? GetInterfaceBuilder(Type collectionType)
    {
        // Create a list or set, and assign it to the member.
        if (typeof(List<TElement>).IsAssignableTo(collectionType))
        {
            return CollectionBuilder<TElement>.ForList();
        }
        if (typeof(HashSet<TElement>).IsAssignableTo(collectionType))
        {
            return CollectionBuilder<TElement>.ForHashSet();
        }

        return null;
    }

    private static CollectionBuilder<TElement>? GetConcreteTypeBuilder(Type collectionType)
    {
        var create = collectionType.GetConstructor([]) is not null ? CompiledAccessors.GetParameterlessConstructor(collectionType) : null;

        // A parameterless constructor and a collection interface.
        if (create is not null && CollectionBuilder<TElement>.ForInterfaces(collectionType, create) is { } builder)
        {
            return builder;
        }

        // A constructor that needs a set, such as ReadOnlySet<T>(ISet<T>).
        if (CollectionUtilities.GetSetConstructor<TElement>(collectionType) is { } setConstructor)
        {
            return CollectionBuilder<TElement>.FromSet(CompiledAccessors.GetConstructor1(setConstructor));
        }

        // A constructor that takes the items as a list, such as ReadOnlyCollection<T>(IList<T>),
        // Queue<T>(IEnumerable<T>) or Stack<T>(IEnumerable<T>). The list is reused if the
        // constructor copies it.
        if (CollectionUtilities.GetListConstructor<TElement>(collectionType) is { } listConstructor)
        {
            var construct = CompiledAccessors.GetConstructor1(listConstructor);
            return CollectionBuilder<TElement>.FromList(construct, CollectionUtilities.ConstructorCopies(collectionType));
        }

        // A parameterless constructor and an Add method.
        if (create is not null && collectionType.GetMethod("Add", [typeof(TElement)]) is { } addMethod)
        {
            return CollectionBuilder<TElement>.Direct(create, CompiledAccessors.GetAction<TElement?>(addMethod));
        }

        // A constructor that takes an ObservableCollection, such as ReadOnlyObservableCollection<T>.
        if (CollectionUtilities.GetConstructorTaking(collectionType, typeof(ObservableCollection<TElement>)) is { } observableConstructor)
        {
            var construct = CompiledAccessors.GetConstructor1(observableConstructor);
            return CollectionBuilder<TElement>.FromList(items => construct(new ObservableCollection<TElement?>(items)), copiesElements: true);
        }

        return null;
    }
}
