﻿using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
{
    public class ChangeTypeMapperTests
    {
        [Theory]
        [InlineData(typeof(string))]
        [InlineData(typeof(int))]
        public void Ctor_Type(Type type)
        {
            var item = new ChangeTypeMapper(type);
            Assert.Equal(type, item.Type);
        }

        [Fact]
        public void Ctor_NullType_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("type", () => new ChangeTypeMapper(null));
        }

        [Fact]
        public void Ctor_TypeNotIConvertible_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("type", () => new ChangeTypeMapper(typeof(List<int>)));
        }

        [Theory]
        [InlineData(typeof(int), "1", 1)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(Type type, string stringValue, object expected)
        {
            var item = new ChangeTypeMapper(type);

            object value = 0;
            PropertyMapperResultType result = item.MapCellValue(new ReadCellValueResult(-1, -1, stringValue), ref value);
            Assert.Equal(PropertyMapperResultType.Success, result);
            Assert.Equal(expected, value);
        }

        [Theory]
        [InlineData(typeof(uint), "-1")]
        [InlineData(typeof(uint), "abc")]
        [InlineData(typeof(uint), "")]
        [InlineData(typeof(uint), null)]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(Type type, string stringValue)
        {
            var item = new ChangeTypeMapper(type);

            object value = 1;
            PropertyMapperResultType result = item.MapCellValue(new ReadCellValueResult(-1, -1, stringValue), ref value);
            Assert.Equal(PropertyMapperResultType.Invalid, result);
            Assert.Equal(1, value);
        }
    }
}
