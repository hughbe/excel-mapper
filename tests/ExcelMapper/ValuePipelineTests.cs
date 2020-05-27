using System;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ValuePipelineTests
    {
        [Fact]
        public void Ctor_Default()
        {
            var pipeline = new ValuePipeline();
            Assert.Empty(pipeline.CellValueMappers);
            Assert.Empty(pipeline.CellValueTransformers);
        }

        [Fact]
        public void EmptyFallback_Set_GetReturnsExpected()
        {
            var pipeline = new ValuePipeline();

            var fallback = new FixedValueFallback(10);
            pipeline.EmptyFallback = fallback;
            Assert.Same(fallback, pipeline.EmptyFallback);

            pipeline.EmptyFallback = null;
            Assert.Null(pipeline.EmptyFallback);
        }

        [Fact]
        public void InvalidFallback_Set_GetReturnsExpected()
        {
            var pipeline = new ValuePipeline();

            var fallback = new FixedValueFallback(10);
            pipeline.InvalidFallback = fallback;
            Assert.Same(fallback, pipeline.InvalidFallback);

            pipeline.InvalidFallback = null;
            Assert.Null(pipeline.InvalidFallback);
        }

        [Fact]
        public void AddCellValueMapper_ValidItem_Success()
        {
            var pipeline = new ValuePipeline();
            var item1 = new BoolMapper();
            var item2 = new BoolMapper();

            pipeline.AddCellValueMapper(item1);
            pipeline.AddCellValueMapper(item2);
            Assert.Equal(new ICellValueMapper[] { item1, item2 }, pipeline.CellValueMappers);
        }

        [Fact]
        public void AddCellValueMapper_NullItem_ThrowsArgumentNullException()
        {
            var pipeline = new ValuePipeline();
            Assert.Throws<ArgumentNullException>("mapper", () => pipeline.AddCellValueMapper(null));
        }

        [Fact]
        public void RemoveCellValueMapper_Index_Success()
        {
            var pipeline = new ValuePipeline();
            pipeline.AddCellValueMapper(new BoolMapper());

            pipeline.RemoveCellValueMapper(0);
            Assert.Empty(pipeline.CellValueMappers);
        }

        [Fact]
        public void AddCellValueTransformer_ValidTransformer_Success()
        {
            var pipeline = new ValuePipeline();
            var transformer1 = new TrimCellValueTransformer();
            var transformer2 = new TrimCellValueTransformer();

            pipeline.AddCellValueTransformer(transformer1);
            pipeline.AddCellValueTransformer(transformer2);
            Assert.Equal(new ICellValueTransformer[] { transformer1, transformer2 }, pipeline.CellValueTransformers);
        }

        [Fact]
        public void AddCellValueTransformer_NullTransformer_ThrowsArgumentNullException()
        {
            var pipeline = new ValuePipeline();
            Assert.Throws<ArgumentNullException>("transformer", () => pipeline.AddCellValueTransformer(null));
        }
    }
}
