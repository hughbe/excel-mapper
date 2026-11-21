namespace ExcelMapper.Utilities;

/// <summary>
/// A collection or dictionary factory that can be copied for one binding (one class map member on
/// one sheet). Factories hold the collection being built, so a factory shared by a map is used under
/// a lock; a binding's own copy is only used by one sheet, so it needs no lock, and can keep buffers
/// between rows.
/// </summary>
internal interface IBindingFactory
{
    /// <summary>
    /// Creates a copy of the factory, with the same configuration and nothing being built.
    /// </summary>
    object CreateForBinding();
}
