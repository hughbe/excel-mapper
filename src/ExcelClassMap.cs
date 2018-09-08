using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper
{
    /// <summary>
    /// A map that maps a row of a sheet to an object of a given type.
    /// </summary>
    public class ExcelClassMap
    {
        /// <summary>
        /// The type of the object to map.
        /// </summary>
        public Type Type { get; }

        public ExcelPropertyMapCollection Mappings { get; } = new ExcelPropertyMapCollection();
    
        /// <summary>
        /// Creates an ExcelClassMap for the given type.
        /// </summary>
        /// <param name="type">The type of the object to map.</param>
        internal ExcelClassMap(Type type) => Type = type;

        /// <summary>
        /// Map the given row of a sheet to an object of a given type. This method goes through each
        /// registered property mapping and uses it to map one or more cells to a property or field
        /// on type of the object to map.
        /// </summary>
        /// <param name="sheet">The sheet that is currently being read.</param>
        /// <param name="rowIndex">The index of the row in the sheet that is currently being read.</param>
        /// <param name="reader">The reader that allows access to the data of the document.</param>
        /// <returns>An object created from one or more cells in the row.</returns>
        internal object Execute(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            object instance = Activator.CreateInstance(Type);

            foreach (ExcelPropertyMap pipeline in Mappings)
            {
                object propertyValue = pipeline.GetPropertyValue(sheet, rowIndex, reader);
                pipeline.SetPropertyFactory(instance, propertyValue);
            }

            return instance;
        }

        /// <summary>
        /// Traverses through a list of member expressions, starting with the member closest to the type
        /// of this class map, and creates a map for each sub member access.
        /// This enables support for expressions such as p => p.prop.subprop.field.final.
        /// </summary>
        /// <param name="propertyMapping">The mapping for the final member access in the stack.</param>
        /// <param name="memberExpressions">A stack of each MemberExpression in the list of member access expressions.</param>
        protected internal void CreateObjectMap(ExcelPropertyMap propertyMapping, Stack<MemberExpression> memberExpressions)
        {
            MemberExpression memberExpression = memberExpressions.Pop();
            if (memberExpressions.Count == 0)
            {
                // This is the final member.
                Mappings.Add(propertyMapping);
                return;
            }

            Type memberType = memberExpression.Member.MemberType();

            MethodInfo method = MapObjectMethod.MakeGenericMethod(memberType);
            try
            {
                method.Invoke(this, new object[] { propertyMapping, memberExpression, memberExpressions });
            }
            catch (TargetInvocationException exception)
            {
                throw exception.InnerException;
            }
        }

        private void CreateObjectMapGeneric<TProperty>(ExcelPropertyMap propertyMapping, MemberExpression memberExpression, Stack<MemberExpression> memberExpressions)
        {
            ExcelPropertyMap mapping = Mappings.FirstOrDefault(m => m.Member.Equals(memberExpression.Member));

            ObjectExcelPropertyMap<TProperty> objectPropertyMapping;
            if (mapping == null)
            {
                objectPropertyMapping = new ObjectExcelPropertyMap<TProperty>(memberExpression.Member, new ExcelClassMap<TProperty>());
                Mappings.Add(objectPropertyMapping);
            }
            else if (!(mapping is ObjectExcelPropertyMap<TProperty> existingMapping))
            {
                throw new InvalidOperationException($"Expression is already mapped differently.");
            }
            else
            {
                objectPropertyMapping = existingMapping;
            }

            objectPropertyMapping.ClassMap.CreateObjectMap(propertyMapping, memberExpressions);
        }

        private static MethodInfo s_mapObjectMethod;
        private static MethodInfo MapObjectMethod => s_mapObjectMethod ?? (s_mapObjectMethod = typeof(ExcelClassMap).GetTypeInfo().GetDeclaredMethod(nameof(CreateObjectMapGeneric)));
    }
}
