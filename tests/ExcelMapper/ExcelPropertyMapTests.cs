using System.Reflection;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;
 
public class PropertyMapTests
{
    [Fact]
    public void Ctor_PropertyInfoMember_Success()
    {
        var member = typeof(PropertyClass).GetProperty(nameof(PropertyClass.Value))!;
        var map = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));

        var propertyMap = new ExcelPropertyMap(member, map);
        Assert.Same(member, propertyMap.Member);
        Assert.Same(map, propertyMap.Map);

        var instance = new PropertyClass();
        propertyMap.SetValueFactory(instance, 10);
        Assert.Equal(10, instance.Value);
    }

    private class PropertyClass
    {
        public int Value { get; set; }
    }

    [Fact]
    public void Ctor_FieldInfoMember_Success()
    {
        MemberInfo fieldInfo = typeof(FieldClass).GetField(nameof(FieldClass.Value))!;
        var map = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));

        var propertyMap = new ExcelPropertyMap(fieldInfo, map);
        Assert.Same(fieldInfo, propertyMap.Member);
        Assert.Same(map, propertyMap.Map);

        var instance = new FieldClass();
        propertyMap.SetValueFactory(instance, 10);
        Assert.Equal(10, instance.Value);
    }

    private class FieldClass
    {
#pragma warning disable CS0649 // Field is never assigned to
        public int Value;
#pragma warning restore CS0649
    }

    [Fact]
    public void Ctor_NullMember_ThrowsArgumentNullException()
    {
        var map = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        Assert.Throws<ArgumentNullException>("member", () => new ExcelPropertyMap(null!, map));
    }

    [Fact]
    public void Ctor_NullMap_ThrowsArgumentNullException()
    {
        var member = typeof(FieldClass).GetField(nameof(FieldClass.Value))!;
        Assert.Throws<ArgumentNullException>("map", () => new ExcelPropertyMap(member, null!));
    }

    [Fact]
    public void Ctor_MemberNotFieldOrProperty_ThrowsArgumentException()
    {
        var member = typeof(EventClass).GetEvent(nameof(EventClass.Event))!;
        var map = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        Assert.Throws<ArgumentException>("member", () => new ExcelPropertyMap(member, map));
    }

    private class EventClass
    {
        public event EventHandler Event { add { } remove { } }
    }

    [Fact]
    public void Ctor_GetOnlyProperty_ThrowsArgumentException()
    {
        var member = typeof(ReadOnlyPropertyClass).GetProperty(nameof(ReadOnlyPropertyClass.ReadOnlyProperty))!;
        var map = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        Assert.Throws<ArgumentException>("member", () => new ExcelPropertyMap(member, map));
    }

    private class ReadOnlyPropertyClass
    {
        public int ReadOnlyProperty { get; } = default!;
    }

    [Fact]
    public void Ctor_StaticProperty_ThrowsArgumentException()
    {
        var propertyInfo = typeof(StaticPropertyClass).GetProperty(nameof(StaticPropertyClass.Value))!;
        var map = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        Assert.Throws<ArgumentException>("member", () => new ExcelPropertyMap(propertyInfo, map));
    }

    private class StaticPropertyClass
    {
        public static int Value { get; set; }
    }

    [Fact]
    public void Ctor_IndexerProperty_ThrowsArgumentException()
    {
        var propertyInfo = typeof(IndexerPropertyClass).GetProperty("Item")!;
        var map = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        Assert.Throws<ArgumentException>("member", () => new ExcelPropertyMap(propertyInfo, map));
    }

    private class IndexerPropertyClass
    {
        public int this[int index]
        {
            get => index;
            set { }
        }
    }
}
