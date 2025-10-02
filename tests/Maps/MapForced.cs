using Xunit;

namespace ExcelMapper.Tests;

public class MapForced
{
    [Fact]
    public void ReadRow_ForcedMappedInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ForcedInt32ValueMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectValue>();
        Assert.Equal(10, row3.Value);
    }

    private class ObjectValue
    {
        public object Value { get; set; } = default!;
    }

    private class ForcedInt32ValueMap : ExcelClassMap<ObjectValue>
    {
        public ForcedInt32ValueMap()
        {
            Map(o => (int)o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }
}
