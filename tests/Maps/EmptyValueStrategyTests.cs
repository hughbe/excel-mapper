namespace ExcelMapper.Tests;

public class EmptyValueStrategyTests
{
    [Fact]
    public void ReadRow_EmptyValueStrategy_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("EmptyValues.xlsx");
        importer.Configuration.RegisterClassMap(new EmptyValueStrategyMap());

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EmptyValues>();
        Assert.Equal(0, row1.IntValue);
        Assert.Null(row1.StringValue);
        Assert.False(row1.BoolValue);
        Assert.Equal((EmptyValuesEnum)0, row1.EnumValue);
        Assert.Equal(DateTime.MinValue, row1.DateValue);
        Assert.Equal([0, 0], row1.ArrayValue);
    }

    public class EmptyValues
    {
        public int IntValue { get; set; }
        public string StringValue { get; set; } = default!;
        public bool BoolValue { get; set; }
        public EmptyValuesEnum EnumValue { get; set; }
        public DateTime DateValue { get; set; }
        public int[] ArrayValue { get; set; } = default!;
    }

    public enum EmptyValuesEnum
    {
        Test = 1
    }

    public class EmptyValueStrategyMap : ExcelClassMap<EmptyValues>
    {
        public EmptyValueStrategyMap() : base(FallbackStrategy.SetToDefaultValue)
        {
            Map(e => e.IntValue);
            Map(e => e.StringValue);
            Map(e => e.BoolValue);
            Map(e => e.EnumValue);
            Map(e => e.DateValue);
            Map(e => e.ArrayValue)
                .WithColumnNames("ArrayValue1", "ArrayValue2");
        }
    }
}
