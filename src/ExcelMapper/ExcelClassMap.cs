using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper
{
    public class ExcelClassMap
    {
        public Type Type { get; }

        internal List<PropertyMapping> Mappings { get; } = new List<PropertyMapping>();

        internal ExcelClassMap(Type type) => Type = type;

        internal object Execute(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            object instance = Activator.CreateInstance(Type);

            foreach (PropertyMapping pipeline in Mappings)
            {
                object propertyValue = pipeline.GetPropertyValue(sheet, rowIndex, reader);
                pipeline.SetPropertyFactory(instance, propertyValue);
            }

            return instance;
        }

        protected PropertyMapping CreateObjectMap(PropertyMapping propertyMapping, Stack<MemberExpression> memberExpressions)
        {
            MemberExpression memberExpression = memberExpressions.Pop();
            if (memberExpressions.Count == 0)
            {
                // This is the final member.
                Mappings.Add(propertyMapping);
                return propertyMapping;
            }

            Type memberType = memberExpression.Member.MemberType();

            MethodInfo method = MapObjectMethod.MakeGenericMethod(memberType);
            try
            {
                return (PropertyMapping)method.Invoke(this, new object[] { propertyMapping, memberExpression, memberExpressions });
            }
            catch (TargetInvocationException exception)
            {
                throw exception.InnerException;
            }
        }

        private PropertyMapping CreateObjectMapGeneric<TProperty>(PropertyMapping propertyMapping, MemberExpression memberExpression, Stack<MemberExpression> memberExpressions)
        {
            PropertyMapping mapping = Mappings.FirstOrDefault(m => m.Member.Equals(memberExpression.Member));

            ObjectPropertyMapping<TProperty> objectPropertyMapping;
            if (mapping == null)
            {
                objectPropertyMapping = new ObjectPropertyMapping<TProperty>(memberExpression.Member, new ExcelClassMap<TProperty>());
                Mappings.Add(objectPropertyMapping);
            }
            else if (!(mapping is ObjectPropertyMapping<TProperty> existingMapping))
            {
                throw new InvalidOperationException($"Expression is already mapped differently.");
            }
            else
            {
                objectPropertyMapping = existingMapping;
            }

            return objectPropertyMapping.ClassMap.CreateObjectMap(propertyMapping, memberExpressions);
        }

        private static MethodInfo s_mapObjectMethod;
        private static MethodInfo MapObjectMethod => s_mapObjectMethod ?? (s_mapObjectMethod = typeof(ExcelClassMap).GetTypeInfo().GetDeclaredMethod(nameof(CreateObjectMapGeneric)));
    }
}
