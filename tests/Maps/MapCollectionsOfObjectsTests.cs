using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// A collection's elements and a dictionary's values are read from cells, so a collection of
/// objects can't be auto-mapped. It's reported with the member and its element type, rather than
/// mapped as an object through the collection's own properties (such as List's Capacity). The
/// source generator reports the same members (EXM008).
/// </summary>
public class MapCollectionsOfObjectsTests
{
    [Fact]
    public void ReadRows_ListOfObjects_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name,Addresses\na,x\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ListOfObjects>().ToArray());
        Assert.Equal("Member \"Addresses\" of type List<MapCollectionsOfObjectsTests.Address> can't be auto-mapped: its elements are MapCollectionsOfObjectsTests.Address, which can't be read from a cell. A collection's elements and a dictionary's values must be values such as strings, numbers, dates or enums. Give MapCollectionsOfObjectsTests.Address a converter ([ExcelConverter] on the member), or ignore the member with [ExcelIgnore].", exception.Message);
    }

    public class ListOfObjects
    {
        public string? Name { get; set; }
        public List<Address>? Addresses { get; set; }
    }

    [Fact]
    public void ReadRows_ArrayOfObjects_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name,Addresses\na,x\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ArrayOfObjects>().ToArray());
        Assert.StartsWith("Member \"Addresses\" of type MapCollectionsOfObjectsTests.Address[] can't be auto-mapped: its elements are MapCollectionsOfObjectsTests.Address,", exception.Message);
    }

    public class ArrayOfObjects
    {
        public string? Name { get; set; }
        public Address[]? Addresses { get; set; }
    }

    [Fact]
    public void ReadRows_DictionaryOfObjects_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name,Home\na,x\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<DictionaryOfObjects>().ToArray());
        Assert.StartsWith("Member \"ByName\" of type Dictionary<string, MapCollectionsOfObjectsTests.Address> can't be auto-mapped: its values are MapCollectionsOfObjectsTests.Address,", exception.Message);
    }

    public class DictionaryOfObjects
    {
        public Dictionary<string, Address>? ByName { get; set; }
    }

    [Fact]
    public void ReadRows_ListOfLists_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Nested\nx\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ListOfLists>().ToArray());
        Assert.StartsWith("Member \"Nested\" of type List<List<string>> can't be auto-mapped: its elements are List<string>,", exception.Message);
    }

    public class ListOfLists
    {
        public List<List<string>>? Nested { get; set; }
    }

    [Fact]
    public void ReadRows_RowsOfListOfObjects_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Street\nx\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<List<Address>>().ToArray());
        Assert.Equal("Type List<MapCollectionsOfObjectsTests.Address> can't be auto-mapped: its elements are MapCollectionsOfObjectsTests.Address, which can't be read from a cell. A collection's elements and a dictionary's values must be values such as strings, numbers, dates or enums. Read rows of a type whose members are read from the columns instead.", exception.Message);
    }

    [Fact]
    public void ReadRows_ListOfObjectsWithConverter_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Addresses\n\"a,b\"\n");
        var sheet = importer.ReadSheet();

        var row = Assert.Single(sheet.ReadRows<ConvertedListOfObjects>());
        Assert.Equal(["a", "b"], row.Addresses!.Select(a => a.Street));
    }

    public class ConvertedListOfObjects
    {
        [ExcelConverter(typeof(AddressConverter))]
        public List<Address>? Addresses { get; set; }
    }

    [Fact]
    public void ReadRows_IgnoredListOfObjects_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var sheet = importer.ReadSheet();

        var row = Assert.Single(sheet.ReadRows<IgnoredListOfObjects>());
        Assert.Equal("a", row.Name);
        Assert.Null(row.Addresses);
    }

    public class IgnoredListOfObjects
    {
        public string? Name { get; set; }

        [ExcelIgnore]
        public List<Address>? Addresses { get; set; }
    }

    [Fact]
    public void ReadRows_NestedMemberCantBeCreated_NamesMember()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<OuterWithInterface>().ToArray());
        Assert.Equal("Cannot auto-map type MapCollectionsOfObjectsTests.OuterWithInterface: member \"Inner.Shape\" of type MapCollectionsOfObjectsTests.IShape is an interface or abstract class, which can't be created.", exception.Message);
    }

    public class OuterWithInterface
    {
        public string? Name { get; set; }
        public InnerWithInterface? Inner { get; set; }
    }

    public class InnerWithInterface
    {
        public IShape? Shape { get; set; }
    }

    public interface IShape
    {
        int Sides { get; set; }
    }

    [Fact]
    public void ReadRows_MemberWithoutConstructor_NamesMember()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<OuterWithoutConstructor>().ToArray());
        Assert.Equal("Cannot auto-map type MapCollectionsOfObjectsTests.OuterWithoutConstructor: member \"Inner\" of type MapCollectionsOfObjectsTests.NoConstructor has neither a public parameterless constructor nor a public constructor whose parameters match its members.", exception.Message);
    }

    public class OuterWithoutConstructor
    {
        public string? Name { get; set; }
        public NoConstructor? Inner { get; set; }
    }

    public class NoConstructor
    {
        public NoConstructor(string other) => Value = other.Length;
        public int Value { get; set; }
    }

    public class Address
    {
        public string? Street { get; set; }
    }

    public class AddressConverter : ICellConverter<Address>
    {
        public CellConversionResult<Address> Convert(CellValue value, CellContext context)
            => CellConversionResult<Address>.Success(new Address { Street = value.Text });
    }
}
