using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Reflection.PortableExecutable;
using ExcelMapper.SourceGeneration;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace ExcelMapper.Tests;

/// <summary>
/// Runs the source generator on code and checks what it writes and reports.
/// </summary>
public class ExcelMappableGeneratorTests
{
    private const string Usings = """
        using System;
        using System.Collections.Generic;
        using ExcelMapper;

        """;

    [Fact]
    public void Generate_Row_RegistersEveryTypeItsMembersUse()
    {
        var (generated, diagnostics) = Generate("""
            namespace App;

            [ExcelMappable(typeof(Employee))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class Employee
            {
                public string Name { get; set; } = "";
                public int? Age { get; set; }
                public DayOfWeek Day { get; set; }
                public DayOfWeek? MaybeDay { get; set; }
                public List<decimal> Amounts { get; set; } = [];
                public string[] Tags { get; set; } = [];
                public Dictionary<string, Guid> Ids { get; set; } = [];
                public Address Address { get; set; } = new();
                public Point Point;
                [ExcelIgnore]
                public TimeSpan Ignored { get; set; }
                public static int Static { get; set; }
                public int this[int index] => index;
            }

            public class Address
            {
                public string? Street { get; set; }
            }

            public struct Point
            {
                public long X { get; set; }
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Equal(
        [
            "Row<global::App.Employee>",
            "Type<string>",
            "Nullable<int>",
            "Type<int>",
            "Enum<global::System.DayOfWeek>",
            "Nullable<global::System.DayOfWeek>",
            "Collection<global::System.Collections.Generic.List<decimal>, decimal>",
            "Type<decimal>",
            "Array<string>",
            "Dictionary<global::System.Collections.Generic.Dictionary<string, global::System.Guid>, string, global::System.Guid>",
            "Type<global::System.Guid>",
            "Type<global::App.Address>",
            "Type<global::App.Point>",
            "Type<long>",
        ], Registrations(generated));
        Assert.Contains("public static AppContext Default { get; } = new();", generated);
        Assert.Contains("public global::ExcelMapper.ExcelClassMap<global::App.Employee> Employee => GetClassMap<global::App.Employee>();", generated);
        Assert.Contains("namespace App", generated);
    }

    [Fact]
    public void Generate_Object_RegistersSettersAndConstructors()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(Person))]
            [ExcelMappable(typeof(Pair))]
            [ExcelMappable(typeof(Point))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class Person : Base
            {
                public Person()
                {
                }

                public Person(string name, int age = 3)
                {
                    Name = name;
                    Age = age;
                }

                protected Person(int age) => Age = age;

                public Person(ref int age) => Age = age;

                public string Name { get; set; } = "";
                public int Age { get; set; }
                public long Field;
                public readonly long ReadOnlyField;
                public string Init { get; init; } = "";
                public string GetOnly => "";
                public int PrivateSet { get; private set; }
                [ExcelIgnore]
                public int Ignored { get; set; }
                public string @class { get; set; } = "";
                public Base? Parent { get; set; }
            }

            public abstract class Base
            {
                protected Base()
                {
                }

                public Base(int inherited) => Inherited = inherited;

                public int Inherited { get; set; }
            }

            public record Pair(string Key, int? Value);

            public struct Point
            {
                public Point(int x) => X = x;

                public int X { get; set; }
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Equal(
        [
            "Constructor<global::Person>(new global::System.Type[] { typeof(string), typeof(int) }, static arguments => new global::Person(arguments.Get<string>(0), arguments.Get<int>(1)))",
            "Setter<global::Person, string>(\"Name\", static (ref global::Person owner, string value) => owner.Name = value)",
            "Setter<global::Person, int>(\"Age\", static (ref global::Person owner, int value) => owner.Age = value)",
            "Setter<global::Person, long>(\"Field\", static (ref global::Person owner, long value) => owner.Field = value)",
            "Setter<global::Person, string>(\"class\", static (ref global::Person owner, string value) => owner.@class = value)",
            "Setter<global::Person, global::Base>(\"Parent\", static (ref global::Person owner, global::Base value) => owner.Parent = value)",
            "Setter<global::Base, int>(\"Inherited\", static (ref global::Base owner, int value) => owner.Inherited = value)",
            "Setter<global::Person, int>(\"Inherited\", static (ref global::Person owner, int value) => owner.Inherited = value)",
            "Constructor<global::Pair>(new global::System.Type[] { typeof(string), typeof(int?) }, static arguments => new global::Pair(arguments.Get<string>(0), arguments.Get<int?>(1)))",
            "Constructor<global::Point>(new global::System.Type[] { typeof(int) }, static arguments => new global::Point(arguments.Get<int>(0)))",
            "Setter<global::Point, int>(\"X\", static (ref global::Point owner, int value) => owner.X = value)",
        ], Accessors(generated));
    }

    [Fact]
    public void Generate_GenericAndArrayRows_NamesProperties()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(Dictionary<string, int>))]
            [ExcelMappable(typeof(int[]))]
            [ExcelMappable(typeof(int?))]
            public partial class AppContext : ExcelMappingContext
            {
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Contains(" DictionaryStringInt32 => ", generated);
        Assert.Contains(" Int32Array => ", generated);
        Assert.Contains(" NullableInt32 => ", generated);
        Assert.DoesNotContain("namespace", generated);
        Assert.Equal(
        [
            "Row<global::System.Collections.Generic.Dictionary<string, int>>",
            "Dictionary<global::System.Collections.Generic.Dictionary<string, int>, string, int>",
            "Type<string>",
            "Type<int>",
            "Row<int[]>",
            "Array<int>",
            "Row<int?>",
            "Nullable<int>",
        ], Registrations(generated));
    }

    [Fact]
    public void Generate_ConversionTypes_RegistersConversion()
    {
        var (generated, diagnostics) = Generate("""
            using System.Collections;
            using System.Diagnostics.CodeAnalysis;
            using System.Numerics;

            [ExcelMappable(typeof(Row))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class Row
            {
                public Code Code { get; set; }
                public Name? Name { get; set; }
                public NoConstructor? NoConstructor { get; set; }
                public Convertible Convertible { get; set; }
                public BigInteger Big { get; set; }
                public ArrayList List { get; set; } = [];
                public Hashtable Table { get; set; } = [];
                public IEnumerable<object> Objects { get; set; } = [];
                public int[,] Grid { get; set; } = new int[0, 0];
            }

            public readonly struct Code : IParsable<Code>
            {
                public static Code Parse(string s, IFormatProvider? provider) => default;
                public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, out Code result) { result = default; return true; }
            }

            public sealed class Name : IParsable<Name>
            {
                public static Name Parse(string s, IFormatProvider? provider) => new();
                public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out Name result) { result = new(); return true; }
            }

            // Not parsed: auto-mapping only parses reference types it can create.
            public sealed class NoConstructor : IParsable<NoConstructor>
            {
                public NoConstructor(string value) => Value = value;
                public string Value { get; }
                public static NoConstructor Parse(string s, IFormatProvider? provider) => new(s);
                public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out NoConstructor result) { result = new(s!); return true; }
            }

            public struct Convertible : IConvertible
            {
                public TypeCode GetTypeCode() => TypeCode.Object;
                public bool ToBoolean(IFormatProvider? provider) => false;
                public byte ToByte(IFormatProvider? provider) => 0;
                public char ToChar(IFormatProvider? provider) => ' ';
                public DateTime ToDateTime(IFormatProvider? provider) => default;
                public decimal ToDecimal(IFormatProvider? provider) => 0;
                public double ToDouble(IFormatProvider? provider) => 0;
                public short ToInt16(IFormatProvider? provider) => 0;
                public int ToInt32(IFormatProvider? provider) => 0;
                public long ToInt64(IFormatProvider? provider) => 0;
                public sbyte ToSByte(IFormatProvider? provider) => 0;
                public float ToSingle(IFormatProvider? provider) => 0;
                public string ToString(IFormatProvider? provider) => "";
                public object ToType(Type conversionType, IFormatProvider? provider) => this;
                public ushort ToUInt16(IFormatProvider? provider) => 0;
                public uint ToUInt32(IFormatProvider? provider) => 0;
                public ulong ToUInt64(IFormatProvider? provider) => 0;
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Equal(
        [
            "Row<global::Row>",
#if LIBRARY_NET8
            "Parsable<global::Code>",
            "ParsableClass<global::Name>",
#else
            // The netstandard builds don't parse IParsable<T> types, so register them as other types.
            "Type<global::Code>",
            "Type<global::Name>",
#endif
            "Type<global::NoConstructor>",
            "Type<string>",
            "Type<global::Convertible>",
            "Type<global::System.Numerics.BigInteger>",
            "Collection<global::System.Collections.ArrayList, object>",
            "Type<object>",
            "Dictionary<global::System.Collections.Hashtable, string, object>",
            "Collection<global::System.Collections.Generic.IEnumerable<object>, object>",
        ], Registrations(generated));
    }

    [Fact]
    public void Generate_AddMethods_FindsElementAndValueTypes()
    {
        var (generated, diagnostics) = Generate("""
            using System.Collections;

            [ExcelMappable(typeof(Row))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class Row
            {
                public Strings Strings { get; set; } = new();
                public Pairs Pairs { get; set; } = new();
            }

            public class Strings : IEnumerable
            {
                public void Add(string value) { }
                public IEnumerator GetEnumerator() => Array.Empty<object>().GetEnumerator();
            }

            public class Pairs : IEnumerable
            {
                public void Add(string key, long value) { }
                public IEnumerator GetEnumerator() => Array.Empty<object>().GetEnumerator();
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Equal(
        [
            "Row<global::Row>",
            "Collection<global::Strings, string>",
            "Type<string>",
            "Dictionary<global::Pairs, string, long>",
            "Type<long>",
        ], Registrations(generated));
    }

    [Fact]
    public void Generate_RecursiveType_RegistersOnce()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(Node))]
            [ExcelMappable(typeof(Node))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class Node
            {
                public Node? Next { get; set; }
                public List<Node> Children { get; set; } = [];
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Equal(["Row<global::Node>", "Collection<global::System.Collections.Generic.List<global::Node>, global::Node>"], Registrations(generated));
        Assert.Single(generated.Split('\n'), line => line.Contains(" Node => "));
    }

    [Fact]
    public void Generate_NestedContext_DeclaresContainingTypes()
    {
        var (generated, diagnostics) = Generate("""
            namespace App.Mapping;

            public partial class Outer
            {
                internal partial struct Inner<T>
                {
                    [ExcelMappable(typeof(int))]
                    internal partial class AppContext : ExcelMappingContext
                    {
                    }
                }
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Contains("partial class Outer", generated);
        Assert.Contains("partial struct Inner<T>", generated);
        Assert.Contains("partial class AppContext", generated);
    }

    [Fact]
    public void Generate_ExistingMembers_KeepsThem()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(Item))]
            [ExcelMappable(typeof(Other.Item))]
            public partial class AppContext : ExcelMappingContext
            {
                public static AppContext Default { get; } = new();
            }

            public class Item
            {
            }

            namespace Other
            {
                public class Item
                {
                }
            }
            """);

        Assert.Empty(diagnostics);
        Assert.DoesNotContain(" Default ", generated);

        // The second type with the same name gets no property.
        Assert.Single(generated.Split('\n'), line => line.Contains(" Item => "));
    }

    [Fact]
    public void Generate_AbstractContext_NoDefault()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(int))]
            public abstract partial class AppContext : ExcelMappingContext
            {
            }
            """);

        Assert.Empty(diagnostics);
        Assert.DoesNotContain(" Default ", generated);
        Assert.Contains("RegisterTypes", generated);
    }

    [Fact]
    public void Generate_ContextWithoutParameterlessConstructor_NoDefault()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(int))]
            public partial class AppContext(string name) : ExcelMappingContext
            {
                public string Name { get; } = name;
            }
            """);

        Assert.Empty(diagnostics);
        Assert.DoesNotContain(" Default ", generated);
    }

    [Fact]
    public void Generate_RecordContext_DeclaresClass()
    {
        var (generated, diagnostics) = Generate("""
            public partial record Outer
            {
                [ExcelMappable(typeof(int))]
                public partial class AppContext : ExcelMappingContext
                {
                }
            }

            public partial record struct OuterStruct
            {
                [ExcelMappable(typeof(int))]
                public partial class AppContext : ExcelMappingContext
                {
                }
            }

            public partial interface IOuter
            {
                [ExcelMappable(typeof(int))]
                public partial class AppContext : ExcelMappingContext
                {
                }
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Contains("partial record Outer", generated);
        Assert.Contains("partial record struct OuterStruct", generated);
        Assert.Contains("partial interface IOuter", generated);
    }

    [Fact]
    public void Generate_NotPartial_ReportsError()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(int))]
            public class AppContext : ExcelMappingContext
            {
                protected override void RegisterTypes(ExcelTypeRegistrar types) { }
            }

            public class Outer
            {
                [ExcelMappable(typeof(int))]
                public partial class NestedContext : ExcelMappingContext
                {
                }
            }
            """, expectErrors: true);

        Assert.Equal(["EXM001", "EXM003", "EXM001"], diagnostics.Select(d => d.Id));
        Assert.Contains("AppContext", diagnostics[0].GetMessage());
        Assert.Equal("", generated);
    }

    [Fact]
    public void Generate_NotContext_ReportsError()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(int))]
            public partial class AppContext
            {
            }
            """);

        Assert.Equal("EXM002", Assert.Single(diagnostics).Id);
        Assert.Equal("", generated);
    }

    [Fact]
    public void Generate_RegisterTypesDeclared_ReportsErrorAndKeepsIt()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(int))]
            public partial class AppContext : ExcelMappingContext
            {
                protected override void RegisterTypes(ExcelTypeRegistrar types) { }
            }
            """);

        Assert.Equal("EXM003", Assert.Single(diagnostics).Id);
        Assert.DoesNotContain("RegisterTypes", generated);
        Assert.Contains(" Int32 => ", generated);
    }

    [Fact]
    public void Generate_PartialAttributes_GeneratesOnce()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(int))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            [ExcelMappable(typeof(string))]
            public partial class AppContext
            {
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Equal(["Row<int>", "Row<string>"], Registrations(generated));
    }

    [Fact]
    public void Generate_OpenGenericType_ReportsError()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(List<>))]
            [ExcelMappable(typeof(int))]
            [Serializable]
            public partial class AppContext : ExcelMappingContext
            {
            }
            """, expectErrors: true);

        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal("EXM004", diagnostic.Id);
        Assert.Contains("List<>", diagnostic.GetMessage());
        Assert.Equal(["Row<int>"], Registrations(generated));
    }

    [Fact]
    public void Generate_AmbiguousAddMethods_RegistersObjects()
    {
        var (generated, diagnostics) = Generate("""
            using System.Collections;

            [ExcelMappable(typeof(Row))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class Row
            {
                public Values Values { get; set; } = new();
            }

            // Two Add methods: the element type is ambiguous, so elements are objects.
            public class Values : IEnumerable
            {
                public void Add(string value) { }
                public void Add(int value) { }
                public IEnumerator GetEnumerator() => Array.Empty<object>().GetEnumerator();
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Equal(["Row<global::Row>", "Collection<global::Values, object>", "Type<object>"], Registrations(generated));
    }

    [Fact]
    public void Generate_UnchangedCompilation_ReusesOutput()
    {
        var compilation = CreateCompilation("""
            [ExcelMappable(typeof(int))]
            public partial class AppContext : ExcelMappingContext
            {
            }
            """);
        GeneratorDriver driver = CSharpGeneratorDriver.Create([new ExcelMappableGenerator().AsSourceGenerator()], parseOptions: s_parseOptions, driverOptions: new GeneratorDriverOptions(IncrementalGeneratorOutputKind.None, trackIncrementalGeneratorSteps: true));
        driver = driver.RunGenerators(compilation);
        driver = driver.RunGenerators(compilation.AddSyntaxTrees(CSharpSyntaxTree.ParseText("class Unrelated { }", s_parseOptions)));

        var outputs = driver.GetRunResult().Results.Single().TrackedOutputSteps.SelectMany(s => s.Value).SelectMany(s => s.Outputs);
        Assert.All(outputs, output => Assert.Contains(output.Reason, new[] { IncrementalStepRunReason.Cached, IncrementalStepRunReason.Unchanged }));
    }

    // The types registered: every registration except setters and constructors.
    private static string[] Registrations(string generated)
        => RegistrationLines(generated).Where(l => !IsAccessor(l)).Select(l => l[..^"()".Length]).ToArray();

    // The setters and constructors registered.
    private static string[] Accessors(string generated)
        => RegistrationLines(generated).Where(IsAccessor).ToArray();

    private static bool IsAccessor(string registration)
        => registration.StartsWith("Setter<", StringComparison.Ordinal) || registration.StartsWith("Constructor<", StringComparison.Ordinal);

    private static IEnumerable<string> RegistrationLines(string generated)
        => generated.Split('\n').Select(l => l.Trim()).Where(l => l.StartsWith("types.", StringComparison.Ordinal)).Select(l => l["types.".Length..^";".Length]);

    private static readonly CSharpParseOptions s_parseOptions = CSharpParseOptions.Default.WithLanguageVersion(LanguageVersion.Preview);

    private static CSharpCompilation CreateCompilation(string source)
        => CSharpCompilation.Create(
            "App",
            [CSharpSyntaxTree.ParseText(Usings + source, s_parseOptions)],
            s_references.Value,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));

    // Runs the generator, and compiles its output, which must have no errors or warnings.
    private static (string Generated, ImmutableArray<Diagnostic> Diagnostics) Generate(string source, bool expectErrors = false)
    {
        var compilation = CreateCompilation(source);
        GeneratorDriver driver = CSharpGeneratorDriver.Create([new ExcelMappableGenerator().AsSourceGenerator()], parseOptions: s_parseOptions);
        driver = driver.RunGeneratorsAndUpdateCompilation(compilation, out var output, out var diagnostics);

        if (!expectErrors && !diagnostics.Any(d => d.Severity == DiagnosticSeverity.Error))
        {
            var problems = output.GetDiagnostics().Where(d => d.Severity >= DiagnosticSeverity.Warning).ToArray();
            Assert.True(problems.Length == 0, string.Join(Environment.NewLine, problems.Select(p => p.ToString())));
        }

        var generated = string.Concat(driver.GetRunResult().GeneratedTrees.Select(t => t.ToString()));
        return (generated, diagnostics);
    }

    private static readonly Lazy<ImmutableArray<MetadataReference>> s_references = new(() =>
    {
        var paths = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!).Split(Path.PathSeparator);
        var references = ImmutableArray.CreateBuilder<MetadataReference>();
        foreach (var path in paths.Append(typeof(ExcelImporter).Assembly.Location).Distinct())
        {
            if (HasMetadata(path))
            {
                references.Add(MetadataReference.CreateFromFile(path));
            }
        }

        return references.ToImmutable();
    });

    private static bool HasMetadata(string path)
    {
        try
        {
            using var stream = File.OpenRead(path);
            using var reader = new PEReader(stream);
            return reader.HasMetadata;
        }
        catch (BadImageFormatException)
        {
            return false;
        }
    }
}
