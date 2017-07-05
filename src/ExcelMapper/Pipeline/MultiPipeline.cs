using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelMapper.Utilities;

namespace ExcelMapper.Pipeline
{
    public abstract class MultiPipeline<T, TElement> : Pipeline<TElement>
    {
        private EnumerableType Type { get; }
        private Func<ICollection<TElement>> Factory { get; }

        internal MultiPipeline(int capacity, MemberInfo member, EmptyValueStrategy emptyValueStrategy) : base(member)
        {
            Type typeofT = typeof(T);
            TypeInfo tInfo = typeof(T).GetTypeInfo();

            if (typeof(Array).GetTypeInfo().IsAssignableFrom(tInfo))
            {
                Type = EnumerableType.Array;
                Factory = () => new List<TElement>(capacity);
            }
            else if (tInfo.IsInterface)
            {
                if (typeof(List<TElement>).ImplementsInterface(typeofT))
                {
                    Type = EnumerableType.Interface;
                    Factory = () => new List<TElement>(capacity);
                }
                else
                {
                    throw new ExcelMappingException($"No known way to create interface \"{tInfo}\". Make sure that \"List<{typeof(TElement)}>\" is assignable from the type.");
                }
            }
            else if (typeof(T).ImplementsInterface(typeof(ICollection<TElement>)))
            {
                Type = EnumerableType.ConcreteType;
                Factory = () => (ICollection<TElement>)Activator.CreateInstance<T>();
            }
            else
            {
                throw new ExcelMappingException();
            }

            AutoMapper.AutoMap(this, emptyValueStrategy);
        }

        protected object CompletePipeline(PipelineContext context, IEnumerable<string> stringValues)
        {
            ICollection<TElement> elements = Factory();

            foreach (string stringValue in stringValues)
            {
                context.StringValue = stringValue;
                TElement element = CompletePipeline(context);
                elements.Add(element);
            }

            if (Type == EnumerableType.Array)
            {
                return elements.ToArray();
            }

            return elements;
        }
    }
}
