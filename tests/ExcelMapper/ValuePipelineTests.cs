using System;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests;

public class ValuePipelineTests
{
    [Fact]
    public void Ctor_Default()
    {
        var pipeline = new ValuePipeline();
        Assert.Empty(pipeline.Mappers);
        Assert.Empty(pipeline.Transformers);
        Assert.Null(pipeline.EmptyFallback);
        Assert.Null(pipeline.InvalidFallback);
    }

    [Fact]
    public void Mappers_Get_ReturnsExpected()
    {
        var pipeline = new ValuePipeline();
        var mappers = pipeline.Mappers;
        Assert.Empty(mappers);
        Assert.Same(mappers, pipeline.Mappers);
    }

    [Fact]
    public void Mappers_AddValidItem_Success()
    {
        var pipeline = new ValuePipeline();
        var item1 = new BoolMapper();
        var item2 = new BoolMapper();

        pipeline.Mappers.Add(item1);
        pipeline.Mappers.Add(item2);
        Assert.Equal([item1, item2], pipeline.Mappers);
    }

    [Fact]
    public void Mappers_AddNullItem_ThrowsArgumentNullException()
    {
        var pipeline = new ValuePipeline();
        Assert.Throws<ArgumentNullException>("item", () => pipeline.Mappers.Add(null!));
    }

    [Fact]
    public void Mappers_RemoveValidItem_Success()
    {
        var pipeline = new ValuePipeline();
        var item1 = new BoolMapper();
        var item2 = new BoolMapper();
        var item3 = new BoolMapper();

        pipeline.Mappers.Add(item1);
        pipeline.Mappers.Add(item2);

        // Remove first.
        pipeline.Mappers.Remove(item1);
        Assert.Equal([item2], pipeline.Mappers);

        // Remove again.
        pipeline.Mappers.Remove(item1);
        Assert.Equal([item2], pipeline.Mappers);

        // Remove non-existing.
        pipeline.Mappers.Remove(item3);
        Assert.Equal([item2], pipeline.Mappers);

        // Remove null.
        pipeline.Mappers.Remove(null!);
        Assert.Equal([item2], pipeline.Mappers);

        // Remove last.
        pipeline.Mappers.Remove(item2);
        Assert.Empty(pipeline.Mappers);
    }

    [Fact]
    public void Mappers_SetValidItem_Success()
    {
        var pipeline = new ValuePipeline();
        var item1 = new BoolMapper();
        var item2 = new BoolMapper();
        pipeline.Mappers.Add(item1);

        pipeline.Mappers[0] = item2;
        Assert.Equal([item2], pipeline.Mappers);
    }

    [Fact]
    public void Mappers_SetNullItem_ThrowsArgumentNullException()
    {
        var pipeline = new ValuePipeline();
        pipeline.Mappers.Add(new BoolMapper());
        Assert.Throws<ArgumentNullException>("item", () => pipeline.Mappers[0] = null!);
    }

    [Fact]
    public void Transformers_Get_ReturnsExpected()
    {
        var pipeline = new ValuePipeline();
        var transformers = pipeline.Transformers;
        Assert.Empty(transformers);
        Assert.Same(transformers, pipeline.Transformers);
    }

    [Fact]
    public void Transformers_AddValidItem_Success()
    {
        var pipeline = new ValuePipeline();
        var item1 = new TrimCellTransformer();
        var item2 = new TrimCellTransformer();

        pipeline.Transformers.Add(item1);
        pipeline.Transformers.Add(item2);
        Assert.Equal([item1, item2], pipeline.Transformers);
    }

    [Fact]
    public void Transformers_AddNullItem_ThrowsArgumentNullException()
    {
        var pipeline = new ValuePipeline();
        Assert.Throws<ArgumentNullException>("item", () => pipeline.Transformers.Add(null!));
    }

    [Fact]
    public void Transformers_RemoveValidItem_Success()
    {
        var pipeline = new ValuePipeline();
        var item1 = new TrimCellTransformer();
        var item2 = new TrimCellTransformer();
        var item3 = new TrimCellTransformer();

        pipeline.Transformers.Add(item1);
        pipeline.Transformers.Add(item2);

        // Remove first.
        pipeline.Transformers.Remove(item1);
        Assert.Equal([item2], pipeline.Transformers);

        // Remove again.
        pipeline.Transformers.Remove(item1);
        Assert.Equal([item2], pipeline.Transformers);

        // Remove non-existing.
        pipeline.Transformers.Remove(item3);
        Assert.Equal([item2], pipeline.Transformers);

        // Remove null.
        pipeline.Transformers.Remove(null!);
        Assert.Equal([item2], pipeline.Transformers);

        // Remove last.
        pipeline.Transformers.Remove(item2);
        Assert.Empty(pipeline.Transformers);
    }

    [Fact]
    public void Transformers_SetValidItem_Success()
    {
        var pipeline = new ValuePipeline();
        var item1 = new TrimCellTransformer();
        var item2 = new TrimCellTransformer();
        pipeline.Transformers.Add(item1);

        pipeline.Transformers[0] = item2;
        Assert.Equal([item2], pipeline.Transformers);
    }

    [Fact]
    public void Transformers_SetNullItem_ThrowsArgumentNullException()
    {
        var pipeline = new ValuePipeline();
        pipeline.Transformers.Add(new TrimCellTransformer());
        Assert.Throws<ArgumentNullException>("item", () => pipeline.Transformers[0] = null!);
    }

    [Fact]
    public void EmptyFallback_Set_GetReturnsExpected()
    {
        var pipeline = new ValuePipeline();

        // Set non-null value.
        var value = new FixedValueFallback(10);
        pipeline.EmptyFallback = value;
        Assert.Same(value, pipeline.EmptyFallback);

        // Set same.
        pipeline.EmptyFallback = value;
        Assert.Same(value, pipeline.EmptyFallback);

        // Set null.
        pipeline.EmptyFallback = null;
        Assert.Null(pipeline.EmptyFallback);
    }

    [Fact]
    public void InvalidFallback_Set_GetReturnsExpected()
    {
        var pipeline = new ValuePipeline();

        // Set non-null value.
        var value = new FixedValueFallback(10);
        pipeline.InvalidFallback = value;
        Assert.Same(value, pipeline.InvalidFallback);

        // Set same.
        pipeline.InvalidFallback = value;
        Assert.Same(value, pipeline.InvalidFallback);

        // Set null.
        pipeline.InvalidFallback = null;
        Assert.Null(pipeline.InvalidFallback);
    }
}
