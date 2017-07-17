using System;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    public delegate void SetPropertyDelegate(object instance, object value);

    /// <summary>
    /// A map that converts one or more values of one or more cells to a value of the
    /// property or field.
    /// </summary>
    public abstract class ExcelPropertyMap
    {
        /// <summary>
        /// The property or field that will be mapped.
        /// </summary>
        public MemberInfo Member { get; }

        /// <summary>
        /// A cached factory that sets the value of a property or field to a given value.
        /// </summary>
        public SetPropertyDelegate SetPropertyFactory { get; }

        /// <summary>
        /// Constructs a map that converts one or more values of one or more cells to a value of
        /// the of property or field.
        /// </summary>
        /// <param name="member">The property or field to map the value of a one or more cells to.</param>
        protected ExcelPropertyMap(MemberInfo member)
        {
            if (member == null)
            {
                throw new ArgumentNullException(nameof(member));
            }

            if (member is PropertyInfo property)
            {
                if (!property.CanWrite)
                {
                    throw new ArgumentException($"Property \"{member.Name}\" is read-only.", nameof(member));
                }

                Member = member;
                SetPropertyFactory = (instance, value) => property.SetValue(instance, value);
            }
            else if (member is FieldInfo field)
            {
                Member = member;
                SetPropertyFactory = (instance, value) => field.SetValue(instance, value);
            }
            else
            {
                throw new ArgumentException($"Member \"{member.Name}\" is not a field or property.", nameof(member));
            }
        }

        /// <summary>
        /// Maps a row of a sheet to an object. The return value will be assigned to the property or field.
        /// </summary>
        /// <param name="sheet">The sheet containing the row that is being read.</param>
        /// <param name="rowIndex">The index of the row that is being read.</param>
        /// <param name="reader">The reader that allows access to the data of the document.</param>
        /// <returns>An object created from one or more cells in the row.</returns>
        public abstract object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}
