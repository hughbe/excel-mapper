using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelMapper.Readers;

/// <summary>
/// The regular expression given to <see cref="ExcelColumnMatchingAttribute"/> or
/// <see cref="ExcelColumnsMatchingAttribute"/>, whose <c>Type</c> is this type.
/// </summary>
internal sealed class RegexColumnMatcher(Regex regex)
{
    public Regex Regex { get; } = regex;
}
