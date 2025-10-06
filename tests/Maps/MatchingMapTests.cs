using System;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Tests;

public class MatchingMapTests
{
    [Fact]
    public void ReadRow_AutoMappedRegexMatching_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ColumnMatchingRegexAttributeClass>();
        Assert.Equal("1", row1.CustomColumn);

        var row2 = sheet.ReadRow<ColumnMatchingRegexAttributeClass>();
        Assert.Equal("3", row2.CustomColumn);

        var row3 = sheet.ReadRow<ColumnMatchingRegexAttributeClass>();
        Assert.Equal("5", row3.CustomColumn);
    }

    [Fact]
    public void ReadRow_AutoMappedRegexNotMatching_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NonColumnMatchingRegexAttributeClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedRegexNotMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NonColumnMatchingRegexOptionalAttributeClass>();
        Assert.Null(row1.CustomColumn);
    }

    [Fact]
    public void ReadRow_AutoMappedMatcherMatching_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ColumnMatchingAttributeClass>();
        Assert.Equal("one", row1.CustomColumn);

        var row2 = sheet.ReadRow<ColumnMatchingAttributeClass>();
        Assert.Equal("First", row2.CustomColumn);

        var row3 = sheet.ReadRow<ColumnMatchingAttributeClass>();
        Assert.Equal("Second", row3.CustomColumn);
    }
    
    [Fact]
    public void ReadRow_AutoMappedMatcherMatchingArguments_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ColumnMatchingArgumentsAttributeClass>();
        Assert.Equal("one", row1.CustomColumn);

        var row2 = sheet.ReadRow<ColumnMatchingArgumentsAttributeClass>();
        Assert.Equal("First", row2.CustomColumn);

        var row3 = sheet.ReadRow<ColumnMatchingArgumentsAttributeClass>();
        Assert.Equal("Second", row3.CustomColumn);
    }

    [Fact]
    public void ReadRow_AutoMappedMatcherNotMatching_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<AlwaysFalseColumnMatchingAttributeClass>());
    }
    
    [Fact]
    public void ReadRow_AutoMappedMatcherNotMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<AlwaysFalseColumnMatchingOptionalAttributeClass>();
        Assert.Null(row1.CustomColumn);
    }

    [Fact]
    public void ReadRow_DefaultMappedRegexMatching_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultColumnMatchingRegexAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ColumnMatchingRegexAttributeClass>();
        Assert.Equal("1", row1.CustomColumn);

        var row2 = sheet.ReadRow<ColumnMatchingRegexAttributeClass>();
        Assert.Equal("3", row2.CustomColumn);

        var row3 = sheet.ReadRow<ColumnMatchingRegexAttributeClass>();
        Assert.Equal("5", row3.CustomColumn);
    }

    [Fact]
    public void ReadRow_DefaultMappedRegexNotMatching_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNonColumnMatchingRegexAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NonColumnMatchingRegexAttributeClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedRegexNotMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("RegexMap.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNonColumnMatchingRegexAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NonColumnMatchingRegexOptionalAttributeClass>();
        Assert.Null(row1.CustomColumn);
    }

    [Fact]
    public void ReadRow_DefaultMappedMatcherMatching_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<DefaultColumnMatchingAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ColumnMatchingAttributeClass>();
        Assert.Equal("one", row1.CustomColumn);

        var row2 = sheet.ReadRow<ColumnMatchingAttributeClass>();
        Assert.Equal("First", row2.CustomColumn);

        var row3 = sheet.ReadRow<ColumnMatchingAttributeClass>();
        Assert.Equal("Second", row3.CustomColumn);
    }
    
    [Fact]
    public void ReadRow_DefaultMappedMatcherMatchingArguments_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<DefaultColumnMatchingArgumentsAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ColumnMatchingArgumentsAttributeClass>();
        Assert.Equal("one", row1.CustomColumn);

        var row2 = sheet.ReadRow<ColumnMatchingArgumentsAttributeClass>();
        Assert.Equal("First", row2.CustomColumn);

        var row3 = sheet.ReadRow<ColumnMatchingArgumentsAttributeClass>();
        Assert.Equal("Second", row3.CustomColumn);
    }

    [Fact]
    public void ReadRow_DefaultMappedMatcherNotMatching_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<DefaultAlwaysFalseColumnMatchingAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<AlwaysFalseColumnMatchingAttributeClass>());
    }
    
    [Fact]
    public void ReadRow_DefaultMappedMatcherNotMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<DefaultAlwaysFalseColumnMatchingOptionalAttributeClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<AlwaysFalseColumnMatchingOptionalAttributeClass>();
        Assert.Null(row1.CustomColumn);
    }

    [Fact]
    public void ReadRow_ColumnNamesHasHeader_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingColumnNamesClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("1", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("2", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("abc", row3.Value);
    }

    [Fact]
    public void ReadRow_ColumnNamesNoHeader_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingColumnNamesClassMap>();

        var sheet = importer.ReadSheet();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_ColumnNamesNothingMatchingNotOptional_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingColumnNamesClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_ColumnNamesNothingMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingColumnNamesOptionalClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Null(row1.Value);
    }

    [Fact]
    public void ReadRow_PredicateHasHeader_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingPredicateClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("1", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("2", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("abc", row3.Value);
    }

    [Fact]
    public void ReadRow_PredicateNoHeader_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingPredicateClassMap>();

        var sheet = importer.ReadSheet();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_PredicateNothingMatchingNotOptional_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingPredicateClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_PredicateNothingMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingPredicateOptionalClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Null(row1.Value);
    }

    [Fact]
    public void ReadRow_MatcherHasHeader_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("1", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("2", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("abc", row3.Value);
    }

    [Fact]
    public void ReadRow_MatcherNoHeader_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<SimpleColumnMatchingClassMap>();

        var sheet = importer.ReadSheet();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("Int Value", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("1", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("2", row3.Value);
    }

    [Fact]
    public void ReadRow_MatcherNothingMatchingNotOptional_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_MatcherNothingMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingOptionalClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Null(row1.Value);
    }

    private class StringValue
    {
        public string Value { get; set; } = default!;
    }

    private class NoColumnMatchingColumnNamesClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingColumnNamesClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching("NoSuchColumn");
        }
    }

    private class NoColumnMatchingColumnNamesOptionalClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingColumnNamesOptionalClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching("NoSuchColumn")
                .MakeOptional();
        }
    }

    private class ColumnMatchingColumnNamesClassMap : ExcelClassMap<StringValue>
    {
        public ColumnMatchingColumnNamesClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching("NoSuchColumn", "Int Value");
        }
    }

    private class ColumnMatchingPredicateClassMap : ExcelClassMap<StringValue>
    {
        public ColumnMatchingPredicateClassMap()
        {
            int i = 0;

            Map(o => o.Value)
                .WithColumnNameMatching(s =>
                {
                    if (s == "Int Value" && i == 0)
                    {
                        i++;
                        return true;
                    }

                    if (s == "StringValue")
                    {
                        i++;
                        return true;
                    }

                    return false;
                });
        }
    }

    private class NoColumnMatchingPredicateClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingPredicateClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching(_ => false);
        }
    }

    private class NoColumnMatchingPredicateOptionalClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingPredicateOptionalClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching(_ => false)
                .MakeOptional();
        }
    }

    private class ColumnMatchingClassMap : ExcelClassMap<StringValue>
    {
        public ColumnMatchingClassMap()
        {
            int i = 0;

            Map(o => o.Value)
                .WithColumnMatching(new ExcelColumnMatcherWithPredicate(s =>
                {
                    if (s == "Int Value" && i == 0)
                    {
                        i++;
                        return true;
                    }

                    if (s == "StringValue")
                    {
                        i++;
                        return true;
                    }

                    return false;
                }));
        }
    }

    private class SimpleColumnMatchingClassMap : ExcelClassMap<StringValue>
    {
        public SimpleColumnMatchingClassMap()
        {
            Map(o => o.Value)
                .WithColumnMatching(new ExcelColumnMatcherWithPredicate(s => true));
        }
    }

    private class NoColumnMatchingClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingClassMap()
        {
            Map(o => o.Value)
                .WithColumnMatching(new ExcelColumnMatcherWithPredicate(_ => false));
        }
    }

    private class NoColumnMatchingOptionalClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingOptionalClassMap()
        {
            Map(o => o.Value)
                .WithColumnMatching(new ExcelColumnMatcherWithPredicate(_ => false))
                .MakeOptional();
        }
    }

    private class ExcelColumnMatcherWithPredicate : IExcelColumnMatcher
    {
        public Func<string, bool> Predicate { get;  }

        public ExcelColumnMatcherWithPredicate(Func<string, bool> predicate)
        {
            Predicate = predicate;
        }

        public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
            => Predicate?.Invoke(sheet.Heading?.GetColumnName(columnIndex) ?? "NoSuchName") ?? true;
    }

    private class ColumnMatchingRegexAttributeClass
    {
        [ExcelColumnMatching(@"Year \d+$")]
        public string CustomColumn { get; set; } = default!;
    }

    private class DefaultColumnMatchingRegexAttributeClassMap : ExcelClassMap<ColumnMatchingRegexAttributeClass>
    {
        public DefaultColumnMatchingRegexAttributeClassMap()
        {
            Map(o => o.CustomColumn);
        }
    }

    private class NonColumnMatchingRegexAttributeClass
    {
        [ExcelColumnMatching("(?!x)x")]
        public string CustomColumn { get; set; } = default!;
    }

    private class DefaultNonColumnMatchingRegexAttributeClassMap : ExcelClassMap<NonColumnMatchingRegexAttributeClass>
    {
        public DefaultNonColumnMatchingRegexAttributeClassMap()
        {
            Map(o => o.CustomColumn);
        }
    }

    private class NonColumnMatchingRegexOptionalAttributeClass
    {
        [ExcelColumnMatching("(?!x)x")]
        [ExcelOptional]
        public string CustomColumn { get; set; } = default!;
    }

    private class DefaultNonColumnMatchingRegexOptionalAttributeClassMap : ExcelClassMap<NonColumnMatchingRegexOptionalAttributeClass>
    {
        public DefaultNonColumnMatchingRegexOptionalAttributeClassMap()
        {
            Map(o => o.CustomColumn);
        }
    }

    private class ColumnMatchingAttributeClass
    {
        [ExcelColumnMatching(typeof(MatchingColumnMatcher))]
        public string CustomColumn { get; set; } = default!;
    }

    private class DefaultColumnMatchingAttributeClassMap : ExcelClassMap<ColumnMatchingAttributeClass>
    {
        public DefaultColumnMatchingAttributeClassMap()
        {
            Map(o => o.CustomColumn);
        }
    }

    private class MatchingColumnMatcher : IExcelColumnMatcher
    {
        public MatchingColumnMatcher()
        {
        }

        public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
            => columnIndex == 1;
    }

    private class ColumnMatchingArgumentsAttributeClass
    {
        [ExcelColumnMatching(typeof(MatchingArgumentsColumnMatcher), ConstructorArguments = new object[] { 1 })]
        public string CustomColumn { get; set; } = default!;
    }

    private class DefaultColumnMatchingArgumentsAttributeClassMap : ExcelClassMap<ColumnMatchingArgumentsAttributeClass>
    {
        public DefaultColumnMatchingArgumentsAttributeClassMap()
        {
            Map(o => o.CustomColumn);
        }
    }

    private class MatchingArgumentsColumnMatcher : IExcelColumnMatcher
    {
        public int ColumnIndex { get; }

        public MatchingArgumentsColumnMatcher(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }

        public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
            => columnIndex == ColumnIndex;
    }

    private class AlwaysFalseColumnMatchingAttributeClass
    {
        [ExcelColumnMatching(typeof(NoMatchingColumnMatcher))]
        public string CustomColumn { get; set; } = default!;
    }

    private class DefaultAlwaysFalseColumnMatchingAttributeClassMap : ExcelClassMap<AlwaysFalseColumnMatchingAttributeClass>
    {
        public DefaultAlwaysFalseColumnMatchingAttributeClassMap()
        {
            Map(o => o.CustomColumn);
        }
    }

    private class AlwaysFalseColumnMatchingOptionalAttributeClass
    {
        [ExcelColumnMatching(typeof(NoMatchingColumnMatcher))]
        [ExcelOptional]
        public string CustomColumn { get; set; } = default!;
    }

    private class DefaultAlwaysFalseColumnMatchingOptionalAttributeClassMap : ExcelClassMap<AlwaysFalseColumnMatchingOptionalAttributeClass>
    {
        public DefaultAlwaysFalseColumnMatchingOptionalAttributeClassMap()
        {
            Map(o => o.CustomColumn);
        }
    }

    private class NoMatchingColumnMatcher : IExcelColumnMatcher
    {
        public NoMatchingColumnMatcher()
        {
        }

        public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
            => false;
    }
}
