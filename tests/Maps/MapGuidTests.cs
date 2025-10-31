namespace ExcelMapper.Tests;

public class MapGuidTests
{
    [Fact]
    public void ReadRow_Guid_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1);

        var row2 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2);

        var row3 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3);

        var row4 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4);

        var row5 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Guid>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Guid>());
    }

    [Fact]
    public void ReadRow_NullableGuid_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1);

        var row2 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2);

        var row3 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3);

        var row4 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4);

        var row5 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5);

        // Empty cell value.
        var row6 = sheet.ReadRow<Guid?>();
        Assert.Null(row6);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Guid?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedGuid_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");
        importer.Configuration.RegisterClassMap<Guid>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1);

        var row2 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2);

        var row3 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3);

        var row4 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4);

        var row5 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Guid>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Guid>());
    }

    [Fact]
    public void ReadRow_CustomMappedGuid_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");
        importer.Configuration.RegisterClassMap<Guid>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f6"))
                .WithInvalidFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f7"));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1);

        var row2 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2);

        var row3 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3);

        var row4 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4);

        var row5 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5);

        // Empty cell value.
        var row6 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f6"), row6);

        // Invalid cell value.
        var row7 = sheet.ReadRow<Guid>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f7"), row7);
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableGuid_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");
        importer.Configuration.RegisterClassMap<Guid?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1);

        var row2 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2);

        var row3 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3);

        var row4 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4);

        var row5 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5);

        // Empty cell value.
        var row6 = sheet.ReadRow<Guid?>();
        Assert.Null(row6);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Guid?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableGuid_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");
        importer.Configuration.RegisterClassMap<Guid?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f6"))
                .WithInvalidFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f7"));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1);

        var row2 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2);

        var row3 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3);

        var row4 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4);

        var row5 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5);

        // Empty cell value.
        var row6 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f6"), row6);

        // Invalid cell value.
        var row7 = sheet.ReadRow<Guid?>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f7"), row7);
    }

    [Fact]
    public void ReadRow_AutoMappedGuidValue_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

        var row2 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

        var row3 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

        var row4 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

        var row5 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedGuidValue_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");
        importer.Configuration.RegisterClassMap<GuidValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

        var row2 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

        var row3 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

        var row4 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

        var row5 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedGuidValue_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");
        importer.Configuration.RegisterClassMap<GuidValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f6"))
                .WithInvalidFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f7"));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

        var row2 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

        var row3 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

        var row4 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

        var row5 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f6"), row6.Value);

        // Invalid cell value.
        var row7 = sheet.ReadRow<GuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f7"), row7.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableGuidValue_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

        var row2 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

        var row3 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

        var row4 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

        var row5 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableGuidValue>();
        Assert.Null(row6.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableGuidValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableGuidValue_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");
        importer.Configuration.RegisterClassMap<NullableGuidValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

        var row2 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

        var row3 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

        var row4 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

        var row5 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableGuidValue>();
        Assert.Null(row6.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableGuidValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableGuidValue_Success()
    {
        using var importer = Helpers.GetImporter("Guids.xlsx");
        importer.Configuration.RegisterClassMap<NullableGuidValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f6"))
                .WithInvalidFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f7"));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

        var row2 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

        var row3 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

        var row4 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

        var row5 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f6"), row6.Value);

        // Invalid cell value.
        var row7 = sheet.ReadRow<NullableGuidValue>();
        Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f7"), row7.Value);
    }

    private class GuidValue
    {
        public Guid Value { get; set; }
    }

    private class NullableGuidValue
    {
        public Guid? Value { get; set; }
    }
}
