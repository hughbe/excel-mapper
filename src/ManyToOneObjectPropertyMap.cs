using System;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper
{
    /// <summary>
    /// A map that reads one or more values from one or more cells and maps these values to each
    /// property and field of the type of the property or field. This is used to map properties
    /// and fields that are objects.
    /// </summary>
    public class ManyToOneObjectPropertyMap<T> : ManyToOnePropertyMap<T>
    {
        private ExcelClassMap<T> _classMap;

        /// <summary>
        /// Gets or sets a class map that maps multiple cells in a row to the properties and fields
        /// of an object.
        /// </summary>
        public ExcelClassMap<T> ClassMap
        {
            get => _classMap;
            set => _classMap = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        /// Constructs a map that reads one or more values from one or more cells and maps these values to each
        /// property and field of the type of the property or field. This is used to map nested objects.
        /// </summary>
        /// <param name="member">The property or field to map the value of a one or more cells to.</param>
        /// <param name="classMap">The class map that maps multiple cells in a row to the properties and fields of an object.</param>
        public ManyToOneObjectPropertyMap(MemberInfo member, ExcelClassMap<T> classMap) : base(member)
        {
            ClassMap = classMap ?? throw new ArgumentNullException(nameof(classMap));
        }

        /// <summary>
        /// Configures the existing class map used to map multiple cells in a row to the properties and fields
        /// of a an object.
        /// </summary>
        /// <param name="classMapFactory">A delegate that allows configuring the default class map used.</param>
        /// <returns>The property map that invoked this method.</returns>
        public ManyToOneObjectPropertyMap<T> WithClassMap(Action<ExcelClassMap<T>> classMapFactory)
        {
            if (classMapFactory == null)
            {
                throw new ArgumentNullException(nameof(classMapFactory));
            }

            classMapFactory(ClassMap);
            return this;
        }

        /// <summary>
        /// Sets the new class map used to map multiple cells in a row to the properties and fields
        /// of a an object.
        /// </summary>
        /// <param name="classMap">The new class map used.</param>
        /// <returns>The property map that invoked this method.</returns>
        public ManyToOneObjectPropertyMap<T> WithClassMap(ExcelClassMap<T> classMap)
        {
            ClassMap = classMap ?? throw new ArgumentNullException(nameof(classMap));
            return this;
        }

        public override void SetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, object instance)
        {
            object result = _classMap.Execute(sheet, rowIndex, reader);
            SetPropertyFactory(instance, result);
        }
    }
}
