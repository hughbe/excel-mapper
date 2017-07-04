using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Pipeline.Items
{
    public class SplitWithDelimiterPipelineItem<T, TElement> : PipelineItem<T>
    {
        public char[] Delimiters { get; private set; }
        public StringSplitOptions Options { get; private set; }

        public SplitWithDelimiterPipelineItem() => Delimiters = new char[] { ',' };

        public SplitWithDelimiterPipelineItem(IEnumerable<char> delimiters)
        {
            if (delimiters == null)
            {
                throw new ArgumentNullException(nameof(delimiters));
            }

            Delimiters = delimiters.ToArray();
        }

        public SplitWithDelimiterPipelineItem<T, TElement> WithNewDelimiters(params char[] delimiters) => WithNewDelimiters((IEnumerable<char>)delimiters);

        public SplitWithDelimiterPipelineItem<T, TElement> WithNewDelimiters(IEnumerable<char> delimiters)
        {
            if (delimiters == null)
            {
                throw new ArgumentNullException(nameof(delimiters));
            }

            Delimiters = delimiters.ToArray();
            return this;
        }

        public SplitWithDelimiterPipelineItem<T, TElement> WithOptions(StringSplitOptions options)
        {
            Options = options;
            return this;
        }

        public override PipelineResult<T> TryMap(PipelineResult<T> item)
        {
            bool isArray = false;
            bool isIEnumerable = false;
            bool isICollection = false;

            if (typeof(Array).GetTypeInfo().IsAssignableFrom(typeof(T).GetTypeInfo()))
            {
                isArray = true;
            }
            else if (typeof(T) == typeof(IEnumerable<TElement>))
            {
                isIEnumerable = true;
            }
            else if (typeof(T) == typeof(ICollection<TElement>))
            {
                isICollection = true;
            }

            if (string.IsNullOrEmpty(item.StringValue))
            {
                T empty;
                if (isArray || isICollection)
                {
                    empty = (T)(object)Array.CreateInstance(typeof(TElement), 0);
                }
                else if (isIEnumerable)
                {
                    empty = (T)Enumerable.Empty<TElement>();
                }
                else
                {
                    empty = Activator.CreateInstance<T>();
                }

                return new PipelineResult<T>(PipelineStatus.Empty, item.StringValue, empty);
            }

            string[] results = item.StringValue.Split(Delimiters, Options);

            ICollection<TElement> collection;
            if (isArray || isIEnumerable || isICollection)
            {
                collection = new List<TElement>(results.Length);
            }
            else
            {
                collection = (ICollection<TElement>)Activator.CreateInstance<T>();
            }

            for (int i = 0; i < results.Length; i++)
            {
                TElement element = (TElement)Convert.ChangeType(results[i], typeof(TElement));
                collection.Add(element);
            }

            if (isArray)
            {
                return item.MakeCompleted((T)(object)collection.ToArray());
            }
            else if (isICollection || isIEnumerable)
            {
                return item.MakeCompleted((T)(object)collection.ToList());
            }

            return item.MakeCompleted((T)collection);
        }
    }
}
