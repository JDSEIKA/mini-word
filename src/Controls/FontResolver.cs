using System;
using System.Collections.Generic;
using System.Drawing;

namespace WordPadClone.Controls
{
    /// <summary>
    /// Chooses an installed font for the scripts present in a run.  GDI+ will
    /// substitute a missing family, but that substitution is not reliable for
    /// every script; this keeps Spanish accents, Cyrillic, Greek, Hebrew,
    /// Arabic, Indic and CJK text readable.
    /// </summary>
    internal static class FontResolver
    {
        private static HashSet<string> _installed;
        private static readonly Dictionary<string, string> Cache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static string Resolve(string family, string text)
        {
            if (string.IsNullOrEmpty(text))
                return Available(family) ? family : FirstAvailable("Segoe UI", "Arial", "Tahoma", "Calibri");

            string script;
            if (Contains(text, 0x0590, 0x05FF)) script = "hebrew";
            else if (Contains(text, 0x0600, 0x06FF) || Contains(text, 0x0750, 0x077F) || Contains(text, 0x08A0, 0x08FF)) script = "arabic";
            else if (Contains(text, 0x0900, 0x0DFF)) script = "indic";
            else if (Contains(text, 0x0E00, 0x0E7F)) script = "thai";
            else if (Contains(text, 0x3040, 0x30FF) || Contains(text, 0x3400, 0x9FFF) ||
                     Contains(text, 0xAC00, 0xD7AF) || Contains(text, 0xF900, 0xFAFF)) script = "cjk";
            else if (Contains(text, 0x1F300, 0x1FAFF)) script = "emoji";
            else script = "latin";

            string key = (family ?? string.Empty) + "|" + script;
            string cached;
            if (Cache.TryGetValue(key, out cached)) return cached;

            string result = family;
            switch (script)
            {
                case "arabic":
                    result = FirstAvailable("Segoe UI", "Tahoma", "Arial", "Ebrima");
                    break;
                case "hebrew":
                    result = FirstAvailable("Segoe UI", "Arial", "Tahoma", "Ebrima");
                    break;
                case "indic":
                    result = FirstAvailable("Nirmala UI", "Segoe UI", "Arial");
                    break;
                case "thai":
                    result = FirstAvailable("Leelawadee UI", "Tahoma", "Segoe UI");
                    break;
                case "cjk":
                    result = FirstAvailable("Microsoft YaHei UI", "Microsoft YaHei", "Yu Gothic UI",
                        "Meiryo", "Malgun Gothic", "SimSun", "Segoe UI");
                    break;
                case "emoji":
                    result = FirstAvailable("Segoe UI Emoji", "Segoe UI Symbol", "Segoe UI");
                    break;
                default:
                    if (!Available(result))
                        result = FirstAvailable("Calibri", "Segoe UI", "Arial", "Tahoma");
                    break;
            }
            if (string.IsNullOrEmpty(result)) result = family;
            Cache[key] = result;
            return result;
        }

        private static bool Contains(string text, int first, int last)
        {
            for (int i = 0; i < text.Length; i++)
            {
                int c = text[i];
                if (c >= first && c <= last) return true;
            }
            return false;
        }

        private static bool Available(string family)
        {
            if (string.IsNullOrEmpty(family)) return false;
            EnsureInstalled();
            return _installed.Contains(family);
        }

        private static string FirstAvailable(params string[] names)
        {
            EnsureInstalled();
            for (int i = 0; i < names.Length; i++)
                if (_installed.Contains(names[i])) return names[i];
            return "Segoe UI";
        }

        private static void EnsureInstalled()
        {
            if (_installed != null) return;
            _installed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                FontFamily[] families = FontFamily.Families;
                for (int i = 0; i < families.Length; i++)
                {
                    _installed.Add(families[i].Name);
                    try { families[i].Dispose(); }
                    catch { }
                }
            }
            catch
            {
                _installed.Add("Segoe UI");
                _installed.Add("Arial");
                _installed.Add("Calibri");
            }
        }
    }
}
