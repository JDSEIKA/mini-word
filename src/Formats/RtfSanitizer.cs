using System.Text.RegularExpressions;

namespace WordPadClone.Formats
{
    /// <summary>
    /// Repairs small RTF conformance problems produced by some generators.
    /// RichEdit stops importing a document if a table row is followed by a
    /// plain paragraph without leaving table mode first.
    /// </summary>
    internal static class RtfSanitizer
    {
        private static readonly Regex RowWithoutPard = new Regex(
            @"(\\row)(\s*\\par)(?!\s*\\pard)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static string Normalize(string rtf)
        {
            if (string.IsNullOrEmpty(rtf)) return rtf;
            return RowWithoutPard.Replace(rtf, "$1\\pard$2");
        }
    }
}
