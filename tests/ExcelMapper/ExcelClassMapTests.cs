using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelClassMapTests : ExcelClassMap<Helpers.TestClass>
    {
        [Fact]
        public void Ctor_Type()
        {
            var map = new ExcelClassMap(typeof(string));
            Assert.Equal(typeof(string), map.Type);
            Assert.Empty(map.Properties);
            Assert.Same(map.Properties, map.Properties);
        }

        [Fact]
        public void Ctor_NullType_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("type", () => new ExcelClassMap(null));
        }
    }
}
