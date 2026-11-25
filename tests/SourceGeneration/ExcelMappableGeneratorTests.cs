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

        // Auto-mapping can't create the abstract Base of Person.Parent, so it can't map Person.
        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal("EXM006", diagnostic.Id);
        Assert.Equal(DiagnosticSeverity.Warning, diagnostic.Severity);
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

            // Parsed by its Parse method, without a parameterless constructor.
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

#if LIBRARY_NET8
        Assert.Empty(diagnostics);
#else
        // The netstandard builds map IParsable<T> types as objects, and Code and Name have no members.
        Assert.Equal(["EXM005: 'Code'", "EXM005: 'Name'"], diagnostics.Select(d => $"{d.Id}: {d.GetMessage().Split(' ')[0]}"));
#endif
        Assert.Equal(
        [
#if LIBRARY_NET8
            "Row<global::Row>",
            "Parsable<global::Code>",
            "ParsableClass<global::Name>",
            "ParsableClass<global::NoConstructor>",
            "Type<global::Convertible>",
            "Type<global::System.Numerics.BigInteger>",
            "Collection<global::System.Collections.ArrayList, object>",
            "Type<object>",
            "Dictionary<global::System.Collections.Hashtable, string, object>",
            "Type<string>",
            "Collection<global::System.Collections.Generic.IEnumerable<object>, object>",
#else
            // The netstandard builds map IParsable<T> types through their members, as objects.
            "Row<global::Row>",
            "Type<global::Code>",
            "Type<global::Name>",
            "Type<global::NoConstructor>",
            "Type<string>",
            "Type<global::Convertible>",
            "Type<global::System.Numerics.BigInteger>",
            "Collection<global::System.Collections.ArrayList, object>",
            "Type<object>",
            "Dictionary<global::System.Collections.Hashtable, string, object>",
            "Collection<global::System.Collections.Generic.IEnumerable<object>, object>",
#endif
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

        // Each row is checked once, as it's registered once. Next contains Node, and Children
        // holds Nodes, which can't be read from a cell.
        Assert.Equal(["EXM007", "EXM008"], diagnostics.Select(d => d.Id));
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
                public int Value { get; set; }
            }

            namespace Other
            {
                public class Item
                {
                    public int Value { get; set; }
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
    public void Generate_TypesAutoMappingCantMap_ReportsWarnings()
    {
        var (generated, diagnostics) = Generate("""
            [ExcelMappable(typeof(Empty))]
            [ExcelMappable(typeof(Shape))]
            [ExcelMappable(typeof(NoConstructor))]
            [ExcelMappable(typeof(Tree))]
            [ExcelMappable(typeof(Fine))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class Empty
            {
                public int PrivateSet { get; private set; }
                public string GetOnly => "";
                [ExcelIgnore]
                public int Ignored { get; set; }
            }

            public interface Shape
            {
                int Sides { get; set; }
            }

            public class NoConstructor
            {
                public NoConstructor(string other) => Other = other;
                public string Other { get; }
                public int Value { get; set; }
            }

            public class Tree
            {
                public Leaf? Leaf { get; set; }
            }

            public class Leaf
            {
                public string? Name { get; set; }
                public Tree? Parent { get; set; }
            }

            public class Fine
            {
                public string? Name { get; set; }
                [ExcelIgnore]
                public Fine? Self { get; set; }
            }
            """);

        Assert.Equal(
        [
            "EXM005: 'Empty' has no members auto-mapping can set, so reading rows of 'Empty' will throw: give it public setters, or a public constructor whose parameters are its properties",
            "EXM006: Auto-mapping can't create 'Shape', so reading rows of 'Shape' will throw: it's an interface or abstract class: map a member of a type that can be created",
            "EXM007: 'Tree' contains itself through member 'Leaf.Parent', so reading rows of 'Tree' will throw: ignore the member with [ExcelIgnore]",
        ], diagnostics.Select(d => $"{d.Id}: {d.GetMessage()}"));
        Assert.All(diagnostics, d => Assert.Equal(DiagnosticSeverity.Warning, d.Severity));

        // The registrations are written as before.
        Assert.Contains("Row<global::Fine>", Registrations(generated));
    }

    [Fact]
    public void Generate_TypeWithoutUsableConstructor_ReportsWarning()
    {
        var (_, diagnostics) = Generate("""
            [ExcelMappable(typeof(NoConstructor))]
            [ExcelMappable(typeof(FromConstructor))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class NoConstructor
            {
                public NoConstructor(string other) { }
                public int Value { get; set; }
            }

            public class FromConstructor
            {
                public FromConstructor(string name, int age = 3) => Name = name;
                public string Name { get; }
            }
            """);

        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal("EXM006", diagnostic.Id);
        Assert.StartsWith("Auto-mapping can't create 'NoConstructor', so reading rows of 'NoConstructor' will throw: it needs a public parameterless constructor", diagnostic.GetMessage());
    }

    [Fact]
    public void Generate_CollectionOfObjects_ReportsWarning()
    {
        var (_, diagnostics) = Generate("""
            [ExcelMappable(typeof(Row))]
            [ExcelMappable(typeof(List<Address>))]
            public partial class AppContext : ExcelMappingContext
            {
            }

            public class Row
            {
                public List<Address>? Addresses { get; set; }
                public Dictionary<string, Address>? ByName { get; set; }
                public List<List<string>>? Nested { get; set; }

                // Elements that are read from a cell.
                public List<int?>? Numbers { get; set; }
                public DayOfWeek[]? Days { get; set; }
                public Dictionary<string, object>? Values { get; set; }

                // A converter reads each element.
                [ExcelConverter(typeof(AddressConverter))]
                public List<Address>? Converted { get; set; }

                [ExcelIgnore]
                public List<Address>? Ignored { get; set; }
            }

            public class Address
            {
                public string? Street { get; set; }
            }

            public class AddressConverter : ICellConverter<Address>
            {
                public CellConversionResult<Address> Convert(CellValue value, CellContext context) => CellConversionResult<Address>.Success(new Address { Street = value.Text });
            }
            """);

        Assert.Equal(
        [
            "EXM008: 'Row.Addresses' holds 'Address', which can't be read from a cell, so reading rows of 'Row' will throw: a collection's elements and a dictionary's values must be values such as strings, numbers or dates; give the member a converter with [ExcelConverter], or ignore it with [ExcelIgnore]",
            "EXM008: 'Row.ByName' holds 'Address', which can't be read from a cell, so reading rows of 'Row' will throw: a collection's elements and a dictionary's values must be values such as strings, numbers or dates; give the member a converter with [ExcelConverter], or ignore it with [ExcelIgnore]",
            "EXM008: 'Row.Nested' holds 'System.Collections.Generic.List<string>', which can't be read from a cell, so reading rows of 'Row' will throw: a collection's elements and a dictionary's values must be values such as strings, numbers or dates; give the member a converter with [ExcelConverter], or ignore it with [ExcelIgnore]",
            "EXM008: 'System.Collections.Generic.List<Address>' holds 'Address', which can't be read from a cell, so reading rows of 'System.Collections.Generic.List<Address>' will throw: a collection's elements and a dictionary's values must be values such as strings, numbers or dates; give the member a converter with [ExcelConverter], or ignore it with [ExcelIgnore]",
        ], diagnostics.Select(d => $"{d.Id}: {d.GetMessage()}"));
        Assert.All(diagnostics, d => Assert.Equal(DiagnosticSeverity.Warning, d.Severity));
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
