using System.Reflection;
using ExcelMapper.Mappings.Support;

namespace ExcelMapper
{
    /// <summary>
    /// Reads a single cell of an excel sheet and maps the value of the cell to the
    /// type T.
    /// </summary>
    /// <typeparam name="T">The type of the member to map the value of a single cell to.</typeparam>
    public class SingleExcelPropertyMap<T> : SingleExcelPropertyMap, ISinglePropertyMapping<T>
    {
        /// <summary>
        /// Constructs a map that reads the value of a single cell and maps the value of the cell
        /// to the type of the property or field.
        /// </summary>
        /// <param name="member">The property or field to map the value of a single cell to.</param>
        public SingleExcelPropertyMap(MemberInfo member) : base(member)
        {
        }
    }
}
