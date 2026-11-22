using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Registering a type creates the generic code auto-mapping needs for it directly, which native AOT
/// compiles (see tests/AotTests). Mapping with the registered types gives the same results as
/// without them. The types here are only used by these tests, so their registrations are the ones
/// auto-mapping uses.
/// </summary>
public class ExcelTypeRegistrarTests
{
    private sealed class Context : ExcelMappingContext
    {
        public static Context Instance { get; } = new();

        protected override void RegisterTypes(ExcelTypeRegistrar types)
        {
            types.Row<Row>();
            types.Type<string>();
            types.Enum<RegistrarColor>();
            types.Nullable<RegistrarColor>();
            types.Type<RegistrarPoint>();
            types.Collection<List<RegistrarListColor>, RegistrarListColor>();
            types.Enum<RegistrarListColor>();
            types.Array<RegistrarArrayColor>();
            types.Enum<RegistrarArrayColor>();
#if LIBRARY_NET8
            types.Parsable<RegistrarCode>();
            types.ParsableClass<RegistrarName>();
#endif

            types.Row<Dictionary<string, RegistrarDictionaryColor>>();
            types.Dictionary<Dictionary<string, RegistrarDictionaryColor>, string, RegistrarDictionaryColor>();
            types.Enum<RegistrarDictionaryColor>();
        }
    }

    public enum RegistrarColor { Red, Green }

    public enum RegistrarListColor { Red, Green }

    public enum RegistrarArrayColor { Red, Green }

    public enum RegistrarDictionaryColor { Red, Green }

    public struct RegistrarPoint
    {
        public int X { get; set; }
        public int Y { get; set; }
    }

    public class GreenConverter : ICellConverter<RegistrarColor>
    {
        public CellConversionResult<RegistrarColor> Convert(CellValue value, CellContext context)
            => value.Text == "g" ? CellConversionResult<RegistrarColor>.Success(RegistrarColor.Green) : CellConversionResult<RegistrarColor>.NotHandled;
    }

    public class Row
    {
        public RegistrarColor Color { get; set; }

        [ExcelConverter(typeof(GreenConverter))]
        public RegistrarColor? Maybe { get; set; }

        public RegistrarPoint Point { get; set; }

        [ExcelSeparators(';')]
        public List<RegistrarListColor> List { get; set; } = [];

        [ExcelSeparators(';')]
        public RegistrarArrayColor[] Array { get; set; } = [];

#if LIBRARY_NET8
        public RegistrarCode Code { get; set; }

        public RegistrarName? Name { get; set; }
#endif
    }

#if LIBRARY_NET8
    public readonly struct RegistrarCode(string value) : IParsable<RegistrarCode>
    {
        public string Value { get; } = value;

        public static RegistrarCode Parse(string s, IFormatProvider? provider) => new(s.ToUpperInvariant());

        public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, out RegistrarCode result)
        {
            result = new RegistrarCode(s?.ToUpperInvariant() ?? "");
            return s is not null;
        }
    }

    // Reference types are only parsed if they have a public parameterless constructor.
    public sealed class RegistrarName : IParsable<RegistrarName>
    {
        public string Value { get; private init; } = "";

        public static RegistrarName Parse(string s, IFormatProvider? provider) => new() { Value = s.ToLowerInvariant() };

        public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out RegistrarName result)
        {
            result = s is null ? null : new RegistrarName { Value = s.ToLowerInvariant() };
            return s is not null;
        }
    }
#endif

    [Fact]
    public void Register_Types_MapsRows()
    {
#if LIBRARY_NET8
        using var importer = Helpers.GetCsvImporter("Color,Maybe,X,Y,List,Array,Code,Name\nGreen,g,1,2,Red;Green,Green,ab,CD\n");
#else
        using var importer = Helpers.GetCsvImporter("Color,Maybe,X,Y,List,Array\nGreen,g,1,2,Red;Green,Green\n");
#endif
        var row = importer.ReadSheet().ReadRows(Context.Instance.GetClassMap<Row>()).Single();
        Assert.Equal(RegistrarColor.Green, row.Color);
        Assert.Equal(RegistrarColor.Green, row.Maybe);
        Assert.Equal(1, row.Point.X);
        Assert.Equal(2, row.Point.Y);
        Assert.Equal([RegistrarListColor.Red, RegistrarListColor.Green], row.List);
        Assert.Equal([RegistrarArrayColor.Green], row.Array);
#if LIBRARY_NET8
        Assert.Equal("AB", row.Code.Value);
        Assert.Equal("cd", row.Name!.Value);
#endif
    }

    [Fact]
    public void Register_Dictionary_MapsRows()
    {
        using var importer = Helpers.GetCsvImporter("A,B\nRed,Green\n");
        var row = importer.ReadSheet().ReadRows(Context.Instance.GetClassMap<Dictionary<string, RegistrarDictionaryColor>>()).Single();
        Assert.Equal(RegistrarDictionaryColor.Red, row["A"]);
        Assert.Equal(RegistrarDictionaryColor.Green, row["B"]);
    }

    // A context whose setters and constructors mark what they set, to show they're the ones used.
    private sealed class AccessorContext : ExcelMappingContext
    {
        public static AccessorContext Instance { get; } = new();

        protected override void RegisterTypes(ExcelTypeRegistrar types)
        {
            types.Row<SetterRow>();
            types.Type<string>();
            types.Type<int>();
            types.Setter<SetterRow, string?>(nameof(SetterRow.Name), static (ref SetterRow row, string? value) => row.Name = value + "!");
            types.Setter<SetterStruct, int>(nameof(SetterStruct.Value), static (ref SetterStruct row, int value) => row.Value = value * 10);
            types.Row<SetterStruct>();

            types.Row<ConstructedRow>();
            types.Nullable<int>();
            types.Dictionary<Dictionary<string, string>, string, string>();
            types.Constructor<ConstructedRow>(
                [typeof(string), typeof(int?), typeof(Dictionary<string, string>), typeof(string), typeof(int)],
                static arguments => new ConstructedRow(arguments.Get<string>(0) + "!", arguments.Get<int?>(1), arguments.Get<Dictionary<string, string>>(2), arguments.Get<string>(3), arguments.Get<int>(4)));
        }
    }

    public class SetterRow
    {
        public string? Name { get; set; }
    }

    public struct SetterStruct
    {
        public int Value { get; set; }
    }

    public record ConstructedRow(
        string Name,
        int? Age,
        [property: ExcelColumnNames("A", "B")] Dictionary<string, string> Values,
        [property: ExcelIgnore] string? Missing = null,
        [property: ExcelIgnore] int Count = 5);

    [Fact]
    public void Setter_Registered_SetsMember()
    {
        using var importer = Helpers.GetCsvImporter("Name\nAl\n");
        Assert.Equal("Al!", importer.ReadSheet().ReadRows(AccessorContext.Instance.GetClassMap<SetterRow>()).Single().Name);
    }

    [Fact]
    public void Setter_RegisteredForStruct_SetsInPlace()
    {
        using var importer = Helpers.GetCsvImporter("Value\n3\n");
        Assert.Equal(30, importer.ReadSheet().ReadRows(AccessorContext.Instance.GetClassMap<SetterStruct>()).Single().Value);
    }

    [Fact]
    public void Constructor_Registered_CreatesWithArguments()
    {
        // Name is read typed, Values as an object (a dictionary), Missing and Count from their
        // default values, and Age is empty.
        using var importer = Helpers.GetCsvImporter("Name,Age,A,B\nAl,,1,2\n");
        var row = importer.ReadSheet().ReadRows(AccessorContext.Instance.GetClassMap<ConstructedRow>()).Single();
        Assert.Equal("Al!", row.Name);
        Assert.Null(row.Age);
        Assert.Equal("1", row.Values["A"]);
        Assert.Equal("2", row.Values["B"]);
        Assert.Null(row.Missing);
        Assert.Equal(5, row.Count);
    }

    [Fact]
    public void Setter_NullArguments_ThrowsArgumentNullException()
    {
        var types = CaptureRegistrar();
        Assert.Throws<ArgumentNullException>("member", () => types.Setter<SetterRow, string>(null!, static (ref SetterRow _, string _) => { }));
        Assert.Throws<ArgumentNullException>("setter", () => types.Setter<SetterRow, string>("Name", null!));
        Assert.Throws<ArgumentNullException>("parameterTypes", () => types.Constructor<SetterRow>(null!, static _ => new SetterRow()));
        Assert.Throws<ArgumentNullException>("constructor", () => types.Constructor<SetterRow>([], null!));
    }

    // The registrar a context is given, which only the library creates.
    private static ExcelTypeRegistrar CaptureRegistrar()
    {
        var context = new CapturingContext();
        context.GetClassMap<string>();
        return context.Types!;
    }

    private sealed class CapturingContext : ExcelMappingContext
    {
        public ExcelTypeRegistrar? Types { get; private set; }

        protected override void RegisterTypes(ExcelTypeRegistrar types)
        {
            Types = types;
            types.Row<string>();
        }
    }
}
