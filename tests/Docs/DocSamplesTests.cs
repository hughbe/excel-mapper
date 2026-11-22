#if LIBRARY_NET8
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace ExcelMapper.Tests;

/// <summary>
/// Compiles every C# sample in README.md and docs/*.md against the library, so the docs can't
/// show an API that doesn't exist or has changed. Each sample is compiled as a C# script, which
/// allows usings, types, members and statements in any order, after <see cref="Prelude"/>, which
/// declares what samples use without declaring it themselves.
/// </summary>
/// <remarks>
/// <list type="bullet">
/// <item>A sample showing 3.x code before 4.0 code marks them with comments starting
/// <c>// 3.x</c> and <c>// 4.0</c>: only the 4.0 code is compiled.</item>
/// <item>A sample that isn't meant to compile (such as an outline with <c>...</c>) is fenced as
/// <c>```csharp no-compile</c>, which renders the same.</item>
/// <item>Samples are only compiled, not run.</item>
/// </list>
/// </remarks>
public class DocSamplesTests
{
    // Names samples use without declaring: the importer, sheet and types most of the docs use.
    private const string Prelude = """
        ExcelImporter importer = null!;
        ExcelSheet sheet = null!;
        Stream stream = null!;
        IExcelDataReader reader = null!;
        string path = "file.xlsx";
        string[] paths = [];
        Dictionary<string, object> row = [];

        public class Employee
        {
            public string Name { get; set; } = "";
            public string Department { get; set; } = "";
            public decimal Salary { get; set; }
            public int Age { get; set; }
            public DateTime BirthDate { get; set; }
            public DateTime HireDate { get; set; }
            public bool Active { get; set; }
            public string Status { get; set; } = "";
            public DateTime[] Certifications { get; set; } = [];
            public float Score { get; set; }
        }

        public class EmployeeMap : ExcelClassMap<Employee>
        {
        }

        public class Product
        {
            public string Name { get; set; } = "";
            public decimal Price { get; set; }
            public int Quantity { get; set; }
        }

        public class Customer
        {
            public string Name { get; set; } = "";
            public Address Address { get; set; } = new();
        }

        public class Address
        {
            public string Street { get; set; } = "";
            public string City { get; set; } = "";
        }

        public class Contact
        {
            public string Phone { get; set; } = "";
        }

        public class Report
        {
            public decimal[] Quarters { get; set; } = [];
        }

        public class Sale
        {
            public string Item { get; set; } = "";
            public decimal Amount { get; set; }
        }

        public class Cost
        {
            public string Item { get; set; } = "";
            public decimal Amount { get; set; }
        }

        // The reader's own code.
        static void Process<T>(T value) { }
        static int ParseWords(string? text) => 0;
        """;

    private static readonly string[] s_usings =
    [
        "System", "System.Collections.Generic", "System.Collections.Immutable", "System.Collections.ObjectModel",
        "System.ComponentModel", "System.ComponentModel.DataAnnotations", "System.Dynamic", "System.Globalization",
        "System.IO", "System.Linq", "System.Text", "System.Text.RegularExpressions", "System.Threading.Tasks",
        "ExcelDataReader", "ExcelMapper",
    ];

    private static readonly CSharpParseOptions s_parseOptions = CSharpParseOptions.Default
        .WithKind(SourceCodeKind.Script)
        .WithLanguageVersion(LanguageVersion.Preview);

    private static readonly CSharpCompilationOptions s_compilationOptions = new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
        .WithUsings(s_usings)
        .WithNullableContextOptions(NullableContextOptions.Disable);

    private static readonly Lazy<ImmutableArray<MetadataReference>> s_references = new(GetReferences);

    private static readonly Lazy<CSharpCompilation> s_prelude = new(() => CSharpCompilation.CreateScriptCompilation(
        "Prelude", CSharpSyntaxTree.ParseText(Prelude, s_parseOptions), s_references.Value, s_compilationOptions));

    [Fact]
    public void Prelude_Compiles()
        => AssertNoErrors(s_prelude.Value, "the prelude");

    [Theory]
    [MemberData(nameof(Samples))]
    public void Sample_Compiles(string file, int line)
    {
        var code = $"#line {line} \"{file}\"\n" + GetSample(file, line);
        var compilation = CSharpCompilation.CreateScriptCompilation(
            "Sample", CSharpSyntaxTree.ParseText(code, s_parseOptions, path: file), s_references.Value, s_compilationOptions,
            previousScriptCompilation: s_prelude.Value);
        AssertNoErrors(compilation, $"the sample at {file}:{line}");
    }

    public static TheoryData<string, int> Samples()
    {
        var data = new TheoryData<string, int>();
        foreach (var file in GetDocFiles())
        {
            var lines = File.ReadAllLines(Path.Combine(RepositoryRoot, file));
            for (int i = 0; i < lines.Length; i++)
            {
                if (IsCSharpFence(lines[i]))
                {
                    // The sample starts on the next line, which is the line number reported.
                    data.Add(file, i + 2);
                    while (++i < lines.Length && !lines[i].TrimStart().StartsWith("```", StringComparison.Ordinal))
                    {
                    }
                }
            }
        }

        return data;
    }

    private static bool IsCSharpFence(string line)
    {
        // A sample fenced as ```csharp no-compile isn't compiled.
        var info = line.Trim();
        return info == "```csharp" || info == "```cs";
    }

    private static string GetSample(string file, int line)
    {
        var lines = File.ReadAllLines(Path.Combine(RepositoryRoot, file));
        var sample = new StringBuilder();
        var skipping = false;
        for (int i = line - 1; i < lines.Length && !lines[i].TrimStart().StartsWith("```", StringComparison.Ordinal); i++)
        {
            // Keep the line count, so errors give the line in the doc.
            var trimmed = lines[i].Trim();
            if (trimmed.StartsWith("// 3.x", StringComparison.Ordinal) || trimmed.StartsWith("// 4.0", StringComparison.Ordinal))
            {
                skipping = trimmed.StartsWith("// 3.x", StringComparison.Ordinal);
                sample.AppendLine();
                continue;
            }

            // C# scripts don't have using declarations: compile "using var x = ..." as "var x = ...".
            var code = trimmed.StartsWith("using var ", StringComparison.Ordinal) ? lines[i].Replace("using var ", "var ") : lines[i];
            sample.AppendLine(skipping ? "" : code);
        }

        return sample.ToString();
    }

    private static void AssertNoErrors(CSharpCompilation compilation, string description)
    {
        var errors = compilation.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Error).ToArray();
        Assert.True(errors.Length == 0, $"Errors compiling {description}:{Environment.NewLine}{string.Join(Environment.NewLine, errors.Select(e => e.ToString()))}");
    }

    private static IEnumerable<string> GetDocFiles()
    {
        yield return "README.md";
        foreach (var file in Directory.GetFiles(Path.Combine(RepositoryRoot, "docs"), "*.md").OrderBy(f => f, StringComparer.Ordinal))
        {
            // The architecture page is for contributors and shows internal code.
            if (Path.GetFileName(file) != "Architecture.md")
            {
                yield return "docs/" + Path.GetFileName(file);
            }
        }
    }

    private static string RepositoryRoot { get; } = FindRepositoryRoot();

    private static string FindRepositoryRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "ExcelMapper.sln")))
            {
                return directory.FullName;
            }
        }

        throw new InvalidOperationException("Can't find the repository root (the directory with ExcelMapper.sln).");
    }

    // The assemblies the tests run with, which include the library and its dependencies.
    private static ImmutableArray<MetadataReference> GetReferences()
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
    }

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
#endif
