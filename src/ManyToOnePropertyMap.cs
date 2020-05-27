using System.Reflection;

namespace ExcelMapper
{
    /// <summary>
    /// Reads multiple cells of an excel sheet and maps the value of the cell to the
    /// type of the property or field.
    /// </summary>
    public abstract class ManyToOnePropertyMap<T> : ExcelPropertyMap
    {
        /// <summary>
        /// Constructs a map that reads one or more values from one or more cells and maps these values to one
        /// property and field of the type of the property or field.
        /// </summary>
        /// <param name="member">The property or field to map the value of a one or more cells to.</param>
        public ManyToOnePropertyMap(MemberInfo member) : base(member)
        {
        }
    }
}
