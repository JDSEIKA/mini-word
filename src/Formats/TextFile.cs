using System;
using System.IO;
using System.Text;

namespace WordPadClone.Formats
{
    public enum TextEncodingKind
    {
        Ansi,
        Utf8,
        Utf8Bom,
        Unicode,
        BigEndianUnicode
    }

    public static class TextFile
    {
        public static string Load(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            string text;

            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                text = new UTF8Encoding(false, true).GetString(bytes, 3, bytes.Length - 3);
            }
            else if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            {
                text = Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            }
            else if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            {
                text = Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            }
            else
            {
                try
                {
                    text = new UTF8Encoding(false, true).GetString(bytes);
                }
                catch (DecoderFallbackException)
                {
                    text = Encoding.Default.GetString(bytes);
                }
            }

            return NormalizeNewlines(text);
        }

        public static void Save(string path, string text, TextEncodingKind kind)
        {
            text = NormalizeNewlines(text).Replace("\n", "\r\n");
            Encoding encoding;
            switch (kind)
            {
                case TextEncodingKind.Ansi:
                    encoding = Encoding.Default;
                    break;
                case TextEncodingKind.Utf8:
                    encoding = new UTF8Encoding(false);
                    break;
                case TextEncodingKind.Utf8Bom:
                    encoding = new UTF8Encoding(true);
                    break;
                case TextEncodingKind.Unicode:
                    encoding = new UnicodeEncoding(false, true);
                    break;
                case TextEncodingKind.BigEndianUnicode:
                    encoding = new UnicodeEncoding(true, true);
                    break;
                default:
                    encoding = new UTF8Encoding(false);
                    break;
            }
            File.WriteAllText(path, text, encoding);
        }

        public static string NormalizeNewlines(string text)
        {
            if (text == null) return string.Empty;
            return text.Replace("\r\n", "\n").Replace('\r', '\n');
        }

    }
}
