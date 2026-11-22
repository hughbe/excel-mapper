using System.Reflection;
using System.Runtime.CompilerServices;

namespace ExcelMapper.Tests;

/// <summary>
/// Makes mapping contexts throw if they map a type they didn't register (the library's internal
/// <c>GenericTypes.RequireRegisteredTypes</c>), as with native AOT, for every test: a generated
/// context whose generator and auto-mapping disagree fails the tests (see "Auto-mapping and the
/// source generator" in CLAUDE.md).
/// </summary>
/// <remarks>
/// The setting is internal, rather than public API or a runtime switch, and the tests don't see the
/// library's internals, so it's set by reflection.
/// </remarks>
internal static class RequireRegisteredTypes
{
    [ModuleInitializer]
    internal static void Enable()
    {
        var property = typeof(ExcelImporter).Assembly.GetType("ExcelMapper.AutoMapping.GenericTypes", throwOnError: true)!
            .GetProperty("RequireRegisteredTypes", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("GenericTypes.RequireRegisteredTypes is missing: the tests need it to check mapping contexts.");
        property.SetValue(null, true);
    }
}
