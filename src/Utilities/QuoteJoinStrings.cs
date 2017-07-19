using System.Collections.Generic;
using System.Linq;

namespace ExcelMapper.Utilities
{
    public static class QuoteJoinStrings
    {
        public static string ArrayJoin(this IEnumerable<string> values)
        {
            IEnumerable<string> quoted = values.Select(v => $"\"{v}\"");
            return $"[{string.Join(", ", quoted)}]";
        }

        public static string ArrayJoin<T>(this IEnumerable<T> values)
        {
            return $"[{string.Join(", ", values)}]";
        }
    }
}
