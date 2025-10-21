﻿using Xunit;

namespace ExcelMapper.Tests;

public class TrimValueTests
{
    [Fact]
    public void ReadRow_CustomMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<TrimStringValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);
    }

    private class StringValue
    {
        public string Value { get; set; } = default!;
    }

    private class TrimStringValueMap : ExcelClassMap<StringValue>
    {
        public TrimStringValueMap()
        {
            Map(o => o.Value)
                .WithTrim();
        }
    }
}
