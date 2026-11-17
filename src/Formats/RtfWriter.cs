using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace WordPadClone.Formats
{
    /// <summary>
    /// Serializes the format-neutral model to RTF.  Raster pictures are emitted
    /// as a private-use placeholder because the RichEdit version shipped with
    /// Windows 7 does not accept \pngblip; the viewer replaces placeholders with
    /// metafiles through the clipboard.
    /// </summary>
    public static class RtfWriter
    {
        public const char ImagePlaceholder = '\uF8FF';

        public static string Write(DocumentModel doc, bool placeholdersForImages)
        {
            List<string> fonts = new List<string>();
            List<Color> colors = new List<Color>();

            AddFont(fonts, "Calibri");
            AddFont(fonts, "Symbol");

            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Paragraph p = doc.Paragraphs[i];
                for (int j = 0; j < p.Inlines.Count; j++)
                {
                    TextInline t = p.Inlines[j] as TextInline;
                    if (t != null)
                    {
                        AddFont(fonts, t.Format.FontFamily);
                        if (t.Format.ForeColor.HasValue) AddColor(colors, t.Format.ForeColor.Value);
                        if (t.Format.BackColor.HasValue) AddColor(colors, t.Format.BackColor.Value);
                    }
                }
            }

            StringBuilder sb = new StringBuilder(4096);
            sb.Append("{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1");
            sb.Append("{\\fonttbl");
            for (int i = 0; i < fonts.Count; i++)
            {
                string cs = string.Equals(fonts[i], "Symbol", StringComparison.OrdinalIgnoreCase) ? "\\fcharset2" : "\\fcharset0";
                sb.Append("{\\f").Append(i).Append("\\fnil").Append(cs).Append(' ').Append(EscapeText(fonts[i])).Append(";}");
            }
            sb.Append('}');

            sb.Append("{\\colortbl ;");
            for (int i = 0; i < colors.Count; i++)
            {
                sb.Append("\\red").Append(colors[i].R)
                  .Append("\\green").Append(colors[i].G)
                  .Append("\\blue").Append(colors[i].B).Append(';');
            }
            sb.Append('}');

            sb.Append("\\paperw").Append(doc.PageWidthTwips)
              .Append("\\paperh").Append(doc.PageHeightTwips)
              .Append("\\margl").Append(doc.MarginLeftTwips)
              .Append("\\margr").Append(doc.MarginRightTwips)
              .Append("\\margt").Append(doc.MarginTopTwips)
              .Append("\\margb").Append(doc.MarginBottomTwips);
            sb.Append("\\viewkind4\\uc1\n");

            if (doc.Paragraphs.Count == 0)
            {
                sb.Append("\\pard\\f0\\fs22\\par\n");
            }

            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Paragraph paragraph = doc.Paragraphs[i];
                if (paragraph.TableRows != null)
                    WriteTable(sb, paragraph);
                else
                    WriteParagraph(sb, paragraph, fonts, colors, placeholdersForImages);
            }

            sb.Append('}');
            return sb.ToString();
        }

        private static void WriteTable(StringBuilder sb, Paragraph p)
        {
            string[][] rows = p.TableRows;
            int columns = 0;
            for (int r = 0; r < rows.Length; r++)
                if (rows[r] != null && rows[r].Length > columns) columns = rows[r].Length;
            if (columns == 0) return;

            int width = p.TableWidthTwips > 0 ? p.TableWidthTwips : 8640;
            int cellWidth = Math.Max(240, width / columns);
            for (int r = 0; r < rows.Length; r++)
            {
                sb.Append("\\trowd\\trgaph108\\trleft0\\trqc");
                for (int c = 0; c < columns; c++)
                {
                    sb.Append("\\clbrdrl\\brdrs\\brdrw15\\clbrdrt\\brdrs\\brdrw15");
                    sb.Append("\\clbrdrr\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15");
                    sb.Append("\\cellx").Append(cellWidth * (c + 1));
                }
                sb.Append("\\intbl ");
                for (int c = 0; c < columns; c++)
                {
                    string cell = rows[r] != null && c < rows[r].Length ? rows[r][c] : string.Empty;
                    if (string.IsNullOrEmpty(cell)) sb.Append("\\par");
                    else AppendEscapedText(sb, cell);
                    sb.Append("\\cell ");
                }
                sb.Append("\\row ");
            }
            sb.Append("\\pard\n");
        }

        private static void WriteParagraph(StringBuilder sb, Paragraph p, List<string> fonts, List<Color> colors, bool images)
        {
            sb.Append("\\pard");
            if (p.PageBreakBefore) sb.Append("\\page ");
            switch (p.Alignment)
            {
                case ParagraphAlignment.Center: sb.Append("\\qc"); break;
                case ParagraphAlignment.Right: sb.Append("\\qr"); break;
                case ParagraphAlignment.Justify: sb.Append("\\qj"); break;
                default: sb.Append("\\ql"); break;
            }

            int left = p.LeftIndentTwips;
            int first = p.FirstLineIndentTwips;
            if (p.Bullet)
            {
                int defaultLeft = 720 * (p.BulletLevel + 1);
                if (left == 0) left = defaultLeft;
                if (first == 0) first = -360;
            }

            sb.Append("\\li").Append(left);
            sb.Append("\\ri").Append(p.RightIndentTwips);
            sb.Append("\\fi").Append(first);
            if (p.SpaceBeforeTwips != 0) sb.Append("\\sb").Append(p.SpaceBeforeTwips);
            if (p.SpaceAfterTwips != 0) sb.Append("\\sa").Append(p.SpaceAfterTwips);
            if (p.LineSpacing == 1) sb.Append("\\sl360\\slmult1");
            else if (p.LineSpacing == 2) sb.Append("\\sl480\\slmult1");

            if (p.TabStopsTwips != null && p.TabStopsTwips.Length > 0)
            {
                sb.Append("\\tabs");
                for (int i = 0; i < p.TabStopsTwips.Length; i++)
                    sb.Append("\\tx").Append(p.TabStopsTwips[i]);
            }

            if (p.Bullet)
            {
                int symbolIndex = fonts.FindIndex(delegate(string s) { return string.Equals(s, "Symbol", StringComparison.OrdinalIgnoreCase); });
                if (symbolIndex < 0) symbolIndex = 0;
                sb.Append("{\\pntext\\f").Append(symbolIndex).Append("\\'B7\\tab}");
                sb.Append("{\\*\\pn\\pnlvlblt\\pnf").Append(symbolIndex).Append("\\pnindent0{\\pntxtb\\'B7}}");
            }

            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline text = p.Inlines[i] as TextInline;
                if (text != null)
                {
                    WriteTextRun(sb, text, fonts, colors);
                    continue;
                }

                ImageInline img = p.Inlines[i] as ImageInline;
                if (img != null && images)
                {
                    sb.Append("\\u").Append(((int)ImagePlaceholder).ToString(CultureInfo.InvariantCulture)).Append('?');
                }
            }

            sb.Append("\\par\n");
        }

        private static void WriteTextRun(StringBuilder sb, TextInline text, List<string> fonts, List<Color> colors)
        {
            RunFormat f = text.Format;
            int fontIndex = fonts.FindIndex(delegate(string s) { return string.Equals(s, f.FontFamily, StringComparison.OrdinalIgnoreCase); });
            if (fontIndex < 0) fontIndex = 0;

            sb.Append("\\f").Append(fontIndex);
            sb.Append("\\fs").Append(f.SizeHalfPoints);
            sb.Append(f.Bold ? "\\b" : "\\b0");
            sb.Append(f.Italic ? "\\i" : "\\i0");
            sb.Append(f.Underline ? "\\ul" : "\\ul0");
            sb.Append(f.Strikeout ? "\\strike" : "\\strike0");
            if (f.ForeColor.HasValue)
                sb.Append("\\cf").Append(ColorIndex(colors, f.ForeColor.Value));
            else
                sb.Append("\\cf0");
            if (f.BackColor.HasValue)
                sb.Append("\\highlight").Append(ColorIndex(colors, f.BackColor.Value));
            else
                sb.Append("\\highlight0");

            if (f.Script > 0) sb.Append("\\super");
            else if (f.Script < 0) sb.Append("\\sub");
            else sb.Append("\\nosupersub");

            sb.Append(' ');
            AppendEscapedText(sb, text.Text);
        }

        private static void AppendEscapedText(StringBuilder sb, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '{': sb.Append("\\{"); break;
                    case '}': sb.Append("\\}"); break;
                    case '\n': sb.Append("\\line "); break;
                    case '\t': sb.Append("\\tab "); break;
                    default:
                        if (c < 32)
                        {
                            sb.Append("\\u").Append(((int)c).ToString(CultureInfo.InvariantCulture)).Append('?');
                        }
                        else if (c < 127)
                        {
                            sb.Append(c);
                        }
                        else
                        {
                            sb.Append("\\u").Append(((int)c).ToString(CultureInfo.InvariantCulture)).Append('?');
                        }
                        break;
                }
            }
        }

        private static string EscapeText(string text)
        {
            StringBuilder sb = new StringBuilder();
            AppendEscapedText(sb, text);
            return sb.ToString();
        }

        private static void AddFont(List<string> fonts, string name)
        {
            if (string.IsNullOrEmpty(name)) name = "Calibri";
            for (int i = 0; i < fonts.Count; i++)
                if (string.Equals(fonts[i], name, StringComparison.OrdinalIgnoreCase)) return;
            fonts.Add(name);
        }

        private static void AddColor(List<Color> colors, Color c)
        {
            for (int i = 0; i < colors.Count; i++)
                if (colors[i].ToArgb() == c.ToArgb()) return;
            colors.Add(c);
        }

        private static int ColorIndex(List<Color> colors, Color c)
        {
            for (int i = 0; i < colors.Count; i++)
                if (colors[i].ToArgb() == c.ToArgb()) return i + 1;
            return 0;
        }

        /// <summary>Inserts page geometry into RTF produced by RichTextBox.</summary>
        public static string ApplyPageSetup(string rtf, DocumentModel pageInfo)
        {
            if (string.IsNullOrEmpty(rtf) || pageInfo == null) return rtf;
            string block = "\\paperw" + pageInfo.PageWidthTwips
                + "\\paperh" + pageInfo.PageHeightTwips
                + "\\margl" + pageInfo.MarginLeftTwips
                + "\\margr" + pageInfo.MarginRightTwips
                + "\\margt" + pageInfo.MarginTopTwips
                + "\\margb" + pageInfo.MarginBottomTwips;

            int index = rtf.IndexOf("\\viewkind", StringComparison.Ordinal);
            if (index < 0) index = rtf.IndexOf("\\pard", StringComparison.Ordinal);
            if (index < 0) return rtf;
            return rtf.Substring(0, index) + block + rtf.Substring(index);
        }
    }
}
