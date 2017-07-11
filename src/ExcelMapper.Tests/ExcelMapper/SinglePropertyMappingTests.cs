using System;
using System.Reflection;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Items;
using ExcelMapper.Mappings.Transformers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class SinglePropertyMappingTests
    {
        [Theory]
        [InlineData(EmptyValueStrategy.SetToDefaultValue)]
        [InlineData(EmptyValueStrategy.ThrowIfPrimitive)]
        public void Ctor_Member_Type_EmptyValueStrategy(EmptyValueStrategy emptyValueStrategy)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            Type type = typeof(int);

            var mapping = new SinglePropertyMapping(propertyInfo, type, emptyValueStrategy);
            Assert.Same(propertyInfo, mapping.Member);
            Assert.Same(type, mapping.Type);

            Assert.Single(mapping.MappingItems);
            Assert.Empty(mapping.StringValueTransformers);
        }

        [Fact]
        public void Ctor_NullType_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            Assert.Throws<ArgumentNullException>("type", () => new SinglePropertyMapping(propertyInfo, null, EmptyValueStrategy.SetToDefaultValue));
        }

        [Theory]
        [InlineData(EmptyValueStrategy.SetToDefaultValue + 1)]
        [InlineData(EmptyValueStrategy.ThrowIfPrimitive - 1)]
        public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentException(EmptyValueStrategy emptyValueStrategy)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            Assert.Throws<ArgumentException>("emptyValueStrategy", () => new SinglePropertyMapping(propertyInfo, typeof(int), emptyValueStrategy));
        }

        [Fact]
        public void EmptyFallback_Set_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SinglePropertyMapping(propertyInfo, typeof(int), EmptyValueStrategy.SetToDefaultValue);

            var fallback = new FixedValueFallback(10);
            mapping.EmptyFallback = fallback;
            Assert.Same(fallback, mapping.EmptyFallback);

            mapping.EmptyFallback = null;
            Assert.Null(mapping.EmptyFallback);
        }

        [Fact]
        public void InvalidFallback_Set_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SinglePropertyMapping(propertyInfo, typeof(int), EmptyValueStrategy.SetToDefaultValue);

            var fallback = new FixedValueFallback(10);
            mapping.InvalidFallback = fallback;
            Assert.Same(fallback, mapping.InvalidFallback);

            mapping.InvalidFallback = null;
            Assert.Null(mapping.InvalidFallback);
        }

        [Fact]
        public void AddMappingItem_ValidItem_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SinglePropertyMapping(propertyInfo, typeof(int), EmptyValueStrategy.SetToDefaultValue);
            ISinglePropertyMappingItem originalItem = Assert.Single(mapping.MappingItems);
            var item1 = new ParseAsBoolMappingItem();
            var item2 = new ParseAsBoolMappingItem();

            mapping.AddMappingItem(item1);
            mapping.AddMappingItem(item2);
            Assert.Equal(new ISinglePropertyMappingItem[] { originalItem, item1, item2 }, mapping.MappingItems);
        }

        [Fact]
        public void AddMappingItem_NullItem_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SinglePropertyMapping(propertyInfo, typeof(int), EmptyValueStrategy.SetToDefaultValue);

            Assert.Throws<ArgumentNullException>("item", () => mapping.AddMappingItem(null));
        }

        [Fact]
        public void RemoveMappingItem_Index_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SinglePropertyMapping(propertyInfo, typeof(int), EmptyValueStrategy.SetToDefaultValue);

            mapping.RemoveMappingItem(0);
            Assert.Empty(mapping.MappingItems);
        }

        [Fact]
        public void AddStringValueTransformer_ValidTransformer_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SinglePropertyMapping(propertyInfo, typeof(int), EmptyValueStrategy.SetToDefaultValue);
            var transformer1 = new TrimStringTransformer();
            var transformer2 = new TrimStringTransformer();

            mapping.AddStringValueTransformer(transformer1);
            mapping.AddStringValueTransformer(transformer2);
            Assert.Equal(new IStringValueTransformer[] { transformer1, transformer2 }, mapping.StringValueTransformers);
        }

        [Fact]
        public void AddStringValueTransformer_NullTransformer_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SinglePropertyMapping(propertyInfo, typeof(int), EmptyValueStrategy.SetToDefaultValue);

            Assert.Throws<ArgumentNullException>("transformer", () => mapping.AddStringValueTransformer(null));
        }

        private class TestClass
        {
            public string Value { get; set; }
        }
    }
}
