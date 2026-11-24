using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace WordPadClone.Formats
{
    public static class DocumentStyles
    {
        public static readonly string[] Names = new string[]
        {
            "Normal", "Título", "Subtítulo", "Título 1", "Título 2", "Título 3", "Cita"
        };

        public static string NameFromId(string id)
        {
            switch ((id ?? string.Empty).ToLowerInvariant())
            {
                case "title": return "Título";
                case "subtitle": return "Subtítulo";
                case "heading1": return "Título 1";
                case "heading2": return "Título 2";
                case "heading3": return "Título 3";
                case "quote": return "Cita";
                default: return "Normal";
            }
        }

        public static string StyleId(string name)
        {
            switch (name)
            {
                case "Título": return "Title";
                case "Subtítulo": return "Subtitle";
                case "Título 1": return "Heading1";
                case "Título 2": return "Heading2";
                case "Título 3": return "Heading3";
                case "Cita": return "Quote";
                default: return "Normal";
            }
        }

        public static void Apply(Paragraph p, string name)
        {
            if (p == null) return;
            p.StyleName = name ?? "Normal";
            string id = StyleId(p.StyleName);
            RunFormat format = new RunFormat("Calibri", 22);
            p.Alignment = ParagraphAlignment.Left;
            p.LeftIndentTwips = 0;
            p.RightIndentTwips = 0;
            p.FirstLineIndentTwips = 0;
            p.SpaceBeforeTwips = 0;
            p.SpaceAfterTwips = 0;
            p.OutlineLevel = 0;
            if (id == "Title")
            {
                format.SizeHalfPoints = 56; format.Bold = true;
                p.Alignment = ParagraphAlignment.Center; p.SpaceAfterTwips = 240;
            }
            else if (id == "Subtitle")
            {
                format.SizeHalfPoints = 32; format.Italic = true; format.ForeColor = Color.FromArgb(89, 89, 89);
                p.Alignment = ParagraphAlignment.Center; p.SpaceAfterTwips = 180;
            }
            else if (id == "Heading1")
            {
                format.SizeHalfPoints = 40; format.Bold = true; p.OutlineLevel = 1;
                p.SpaceBeforeTwips = 240; p.SpaceAfterTwips = 120; p.KeepWithNext = true;
            }
            else if (id == "Heading2")
            {
                format.SizeHalfPoints = 32; format.Bold = true; p.OutlineLevel = 2;
                p.SpaceBeforeTwips = 200; p.SpaceAfterTwips = 100; p.KeepWithNext = true;
            }
            else if (id == "Heading3")
            {
                format.SizeHalfPoints = 28; format.Bold = true; format.Italic = true; p.OutlineLevel = 3;
                p.SpaceBeforeTwips = 160; p.SpaceAfterTwips = 80; p.KeepWithNext = true;
            }
            else if (id == "Quote")
            {
                format.SizeHalfPoints = 22; format.Italic = true; format.ForeColor = Color.FromArgb(89, 89, 89);
                p.LeftIndentTwips = 720; p.RightIndentTwips = 720; p.SpaceBeforeTwips = 120; p.SpaceAfterTwips = 120;
            }
            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline run = p.Inlines[i] as TextInline;
                if (run == null) continue;
                if (run.Format == null) run.Format = format.Clone();
                else
                {
                    run.Format.FontFamily = format.FontFamily;
                    run.Format.SizeHalfPoints = format.SizeHalfPoints;
                    run.Format.Bold = format.Bold;
                    run.Format.Italic = format.Italic;
                    run.Format.ForeColor = format.ForeColor;
                }
            }
            if (p.Inlines.Count == 0) p.Inlines.Add(new TextInline(string.Empty, format));
        }
    }
    /// <summary>Fields, revision marks, document comparison and TOC helpers.</summary>
    public static class FieldText
    {
        public static FieldKind ParseInstruction(string instruction)
        {
            string upper = (instruction ?? string.Empty).Trim().ToUpperInvariant();
            if (upper.StartsWith("PAGEREF")) return FieldKind.PageRef;
            if (upper.StartsWith("PAGE")) return FieldKind.Page;
            if (upper.StartsWith("NUMPAGES") || upper.StartsWith("SECTIONPAGES")) return FieldKind.NumPages;
            if (upper.StartsWith("DATE") || upper.StartsWith("CREATEDATE")
                || upper.StartsWith("SAVEDATE") || upper.StartsWith("PRINTDATE")) return FieldKind.Date;
            if (upper.StartsWith("TIME")) return FieldKind.Time;
            if (upper.StartsWith("TITLE")) return FieldKind.Title;
            if (upper.StartsWith("AUTHOR")) return FieldKind.Author;
            if (upper.StartsWith("FILENAME")) return FieldKind.FileName;
            if (upper.StartsWith("DOCVARIABLE")) return FieldKind.Custom;
            return FieldKind.None;
        }

        public static string Instruction(FieldKind kind, string custom)
        {
            switch (kind)
            {
                case FieldKind.Page: return " PAGE ";
                case FieldKind.NumPages: return " NUMPAGES ";
                case FieldKind.Date: return " DATE ";
                case FieldKind.Time: return " TIME ";
                case FieldKind.Title: return " TITLE ";
                case FieldKind.Author: return " AUTHOR ";
                case FieldKind.FileName: return " FILENAME ";
                case FieldKind.PageRef: return " PAGEREF " + (custom ?? string.Empty) + " \\h ";
                default: return custom ?? string.Empty;
            }
        }

        public static string Display(FieldKind kind, DocumentModel doc, int page, int pageCount)
        {
            switch (kind)
            {
                case FieldKind.Page: return Math.Max(1, page).ToString(CultureInfo.InvariantCulture);
                case FieldKind.NumPages: return Math.Max(1, pageCount).ToString(CultureInfo.InvariantCulture);
                case FieldKind.Date: return DateTime.Now.ToString("d 'de' MMMM 'de' yyyy", CultureInfo.CurrentCulture);
                case FieldKind.Time: return DateTime.Now.ToString("t", CultureInfo.CurrentCulture);
                case FieldKind.Title: return doc == null ? string.Empty : (doc.Title ?? string.Empty);
                case FieldKind.Author: return doc == null ? string.Empty : (doc.Author ?? string.Empty);
                case FieldKind.FileName: return string.Empty;
                default: return string.Empty;
            }
        }

        public static TextInline Create(FieldKind kind, DocumentModel doc, int page, int pageCount)
        {
            string instruction = Instruction(kind, null);
            TextInline run = new TextInline(Display(kind, doc, page, pageCount), new RunFormat("Calibri", 22));
            run.Field = kind;
            run.FieldInstruction = instruction;
            return run;
        }
    }

    public static class DocumentRevisions
    {
        public static void MarkInsert(TextInline run)
        {
            if (run == null) return;
            run.Revision = RevisionKind.Insert;
            run.RevisionAuthor = "Mini Word";
            run.RevisionDate = DateTime.Now.ToString("s", CultureInfo.InvariantCulture);
        }

        public static void MarkDelete(TextInline run)
        {
            if (run == null) return;
            run.Revision = RevisionKind.Delete;
            run.RevisionAuthor = "Mini Word";
            run.RevisionDate = DateTime.Now.ToString("s", CultureInfo.InvariantCulture);
        }

        public static int Count(DocumentModel doc)
        {
            int count = 0;
            if (doc == null) return 0;
            for (int i = 0; i < doc.Paragraphs.Count; i++)
                for (int j = 0; j < doc.Paragraphs[i].Inlines.Count; j++)
                {
                    TextInline run = doc.Paragraphs[i].Inlines[j] as TextInline;
                    if (run != null && run.Revision != RevisionKind.None) count++;
                }
            return count;
        }

        public static bool AcceptAll(DocumentModel doc)
        {
            if (doc == null || !HasRevisions(doc)) return false;
            for (int i = doc.Paragraphs.Count - 1; i >= 0; i--)
            {
                Paragraph p = doc.Paragraphs[i];
                for (int j = p.Inlines.Count - 1; j >= 0; j--)
                {
                    TextInline run = p.Inlines[j] as TextInline;
                    if (run == null || run.Revision == RevisionKind.None) continue;
                    if (run.Revision == RevisionKind.Delete) p.Inlines.RemoveAt(j);
                    else run.Revision = RevisionKind.None;
                }
                if (p.Inlines.Count == 0 && p.TableRows == null)
                    p.Inlines.Add(new TextInline(string.Empty, new RunFormat()));
            }
            return true;
        }

        public static bool RejectAll(DocumentModel doc)
        {
            if (doc == null || !HasRevisions(doc)) return false;
            for (int i = doc.Paragraphs.Count - 1; i >= 0; i--)
            {
                Paragraph p = doc.Paragraphs[i];
                for (int j = p.Inlines.Count - 1; j >= 0; j--)
                {
                    TextInline run = p.Inlines[j] as TextInline;
                    if (run == null || run.Revision == RevisionKind.None) continue;
                    if (run.Revision == RevisionKind.Insert) p.Inlines.RemoveAt(j);
                    else run.Revision = RevisionKind.None;
                }
                if (p.Inlines.Count == 0 && p.TableRows == null)
                    p.Inlines.Add(new TextInline(string.Empty, new RunFormat()));
            }
            return true;
        }

        public static bool AcceptAt(Paragraph p, int offset)
        {
            TextInline run = RunAt(p, offset);
            if (run == null || run.Revision == RevisionKind.None) return false;
            if (run.Revision == RevisionKind.Delete)
            {
                p.Inlines.Remove(run);
                if (p.Inlines.Count == 0 && p.TableRows == null)
                    p.Inlines.Add(new TextInline(string.Empty, new RunFormat()));
            }
            else run.Revision = RevisionKind.None;
            return true;
        }

        public static bool RejectAt(Paragraph p, int offset)
        {
            TextInline run = RunAt(p, offset);
            if (run == null || run.Revision == RevisionKind.None) return false;
            if (run.Revision == RevisionKind.Insert)
            {
                p.Inlines.Remove(run);
                if (p.Inlines.Count == 0 && p.TableRows == null)
                    p.Inlines.Add(new TextInline(string.Empty, new RunFormat()));
            }
            else run.Revision = RevisionKind.None;
            return true;
        }

        private static TextInline RunAt(Paragraph p, int offset)
        {
            if (p == null || offset < 0) return null;
            int position = 0;
            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline run = p.Inlines[i] as TextInline;
                if (run == null) continue;
                if (offset >= position && offset <= position + run.Text.Length)
                    return run;
                position += run.Text.Length;
            }
            return null;
        }
        private static bool HasRevisions(DocumentModel doc)
        {
            for (int i = 0; i < doc.Paragraphs.Count; i++)
                for (int j = 0; j < doc.Paragraphs[i].Inlines.Count; j++)
                {
                    TextInline run = doc.Paragraphs[i].Inlines[j] as TextInline;
                    if (run != null && run.Revision != RevisionKind.None) return true;
                }
            return false;
        }

        /// <summary>Marks the selected range as deleted without removing it.</summary>
        public static void MarkRange(Paragraph p, int start, int end)
        {
            if (p == null || start >= end) return;
            int offset = 0;
            List<Inline> result = new List<Inline>();
            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline run = p.Inlines[i] as TextInline;
                if (run == null)
                {
                    result.Add(p.Inlines[i]);
                    continue;
                }
                int length = run.Text.Length;
                int from = Math.Max(start, offset);
                int to = Math.Min(end, offset + length);
                if (to <= from)
                {
                    result.Add(run);
                }
                else
                {
                    if (from > offset) result.Add(new TextInline(run.Text.Substring(0, from - offset), run.Format.Clone()));
                    TextInline deleted = new TextInline(run.Text.Substring(from - offset, to - from), run.Format.Clone());
                    MarkDelete(deleted);
                    result.Add(deleted);
                    if (to < offset + length) result.Add(new TextInline(run.Text.Substring(to - offset), run.Format.Clone()));
                }
                offset += length;
            }
            p.Inlines.Clear();
            p.Inlines.AddRange(result);
        }
    }

    public static class TableOfContents
    {
        public static List<Paragraph> Build(DocumentModel doc, int levels, Func<int, int> pageOf)
        {
            List<Paragraph> result = new List<Paragraph>();
            if (doc == null) return result;
            int maxLevel = Math.Max(1, Math.Min(6, levels));
            Paragraph title = new Paragraph();
            title.IsTocTitle = true;
            title.SpaceAfterTwips = 120;
            RunFormat titleFormat = new RunFormat("Calibri", 28);
            titleFormat.Bold = true;
            title.Inlines.Add(new TextInline("Índice", titleFormat));
            result.Add(title);

            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Paragraph p = doc.Paragraphs[i];
                if (p.IsTocEntry || p.IsTocTitle || p.TableRows != null) continue;
                int level = p.OutlineLevel;
                if (level <= 0 || level > maxLevel) continue;
                string text = ParagraphText(p).Trim();
                if (text.Length == 0) continue;
                Paragraph entry = new Paragraph();
                entry.IsTocEntry = true;
                entry.TocLevel = level;
                entry.TocTargetParagraph = i;
                entry.TocPage = pageOf == null ? 0 : pageOf(i);
                entry.LeftIndentTwips = (level - 1) * 360;
                entry.TabStopsTwips = new int[] { 9000 };
                RunFormat format = new RunFormat("Calibri", 22);
                entry.Inlines.Add(new TextInline(text + "\t", format));
                entry.Inlines.Add(new TextInline(entry.TocPage > 0
                    ? entry.TocPage.ToString(CultureInfo.InvariantCulture) : string.Empty, format));
                result.Add(entry);
            }
            return result;
        }

        public static int RemoveExisting(DocumentModel doc, int at)
        {
            if (doc == null) return at;
            int index = Math.Max(0, Math.Min(doc.Paragraphs.Count, at));
            if (index > 0 && doc.Paragraphs[index - 1].IsTocEntry)
            {
                while (index > 0 && doc.Paragraphs[index - 1].IsTocEntry) index--;
            }
            if (index > 0 && doc.Paragraphs[index - 1].IsTocTitle) index--;
            while (index < doc.Paragraphs.Count && (doc.Paragraphs[index].IsTocEntry || doc.Paragraphs[index].IsTocTitle))
                doc.Paragraphs.RemoveAt(index);
            return index;
        }

        public static bool RefreshPages(DocumentModel doc, Func<int, int> pageOf)
        {
            if (doc == null || pageOf == null) return false;
            bool changed = false;
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Paragraph p = doc.Paragraphs[i];
                if (!p.IsTocEntry || p.TocTargetParagraph < 0) continue;
                int page = pageOf(p.TocTargetParagraph);
                if (page <= 0 || page == p.TocPage) continue;
                p.TocPage = page;
                if (p.Inlines.Count > 0)
                {
                    TextInline last = p.Inlines[p.Inlines.Count - 1] as TextInline;
                    if (last != null) last.Text = page.ToString(CultureInfo.InvariantCulture);
                }
                changed = true;
            }
            return changed;
        }
        private static string ParagraphText(Paragraph p)
        {
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline run = p.Inlines[i] as TextInline;
                if (run != null) text.Append(run.Text);
            }
            return text.ToString();
        }
    }

    public static class DocumentCompare
    {
        public static DocumentModel Compare(DocumentModel oldDoc, DocumentModel newDoc)
        {
            return Compare(oldDoc, newDoc, null);
        }

        public static DocumentModel Compare(DocumentModel oldDoc, DocumentModel newDoc, List<string> formatNotes)
        {
            if (newDoc == null) return oldDoc;
            DocumentModel result = new DocumentModel();
            CopyDocumentProperties(newDoc, result);
            if (oldDoc == null)
            {
                for (int i = 0; i < newDoc.Paragraphs.Count; i++)
                    result.Paragraphs.Add(CloneParagraph(newDoc.Paragraphs[i], RevisionKind.Insert));
                return result;
            }

            string[] a = new string[oldDoc.Paragraphs.Count];
            string[] b = new string[newDoc.Paragraphs.Count];
            for (int i = 0; i < a.Length; i++) a[i] = PlainText(oldDoc.Paragraphs[i]);
            for (int i = 0; i < b.Length; i++) b[i] = PlainText(newDoc.Paragraphs[i]);
            List<int[]> ops = Diff(a, b);
            for (int i = 0; i < ops.Count; i++)
            {
                int kind = ops[i][0], oi = ops[i][1], ni = ops[i][2];
                if (kind == 0)
                {
                    if (formatNotes != null && !string.Equals(FormatKey(oldDoc.Paragraphs[oi]), FormatKey(newDoc.Paragraphs[ni]), StringComparison.Ordinal))
                        formatNotes.Add("Párrafo " + (oi + 1) + ": cambió el formato");
                    result.Paragraphs.Add(DiffParagraph(oldDoc.Paragraphs[oi], newDoc.Paragraphs[ni]));
                }
                else if (kind == 1)
                {
                    // Pair a removed paragraph with the next inserted one so a
                    // one-word edit is shown as a revision instead of replacing
                    // the whole paragraph.
                    if (i + 1 < ops.Count && ops[i + 1][0] == 2)
                    {
                        result.Paragraphs.Add(DiffParagraph(oldDoc.Paragraphs[oi], newDoc.Paragraphs[ops[i + 1][2]]));
                        i++;
                    }
                    else
                    {
                        result.Paragraphs.Add(CloneParagraph(oldDoc.Paragraphs[oi], RevisionKind.Delete));
                    }
                }
                else
                {
                    result.Paragraphs.Add(CloneParagraph(newDoc.Paragraphs[ni], RevisionKind.Insert));
                }
            }
            if (result.Paragraphs.Count == 0) result.Paragraphs.Add(new Paragraph());
            result.TrackChanges = true;
            result.ShowRevisions = true;
            return result;
        }

        private static string FormatKey(Paragraph p)
        {
            if (p == null) return string.Empty;
            StringBuilder key = new StringBuilder();
            key.Append((int)p.Alignment).Append('|').Append(p.LeftIndentTwips).Append('|')
                .Append(p.RightIndentTwips).Append('|').Append(p.FirstLineIndentTwips).Append('|')
                .Append(p.SpaceBeforeTwips).Append('|').Append(p.SpaceAfterTwips).Append('|')
                .Append(p.LineSpacing).Append('|').Append(p.OutlineLevel).Append('|')
                .Append(p.StyleName).Append('|').Append(p.Bullet).Append(p.BulletLevel)
                .Append(p.Numbered).Append(p.NumberLevel);
            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline run = p.Inlines[i] as TextInline;
                if (run == null || run.Format == null) continue;
                RunFormat f = run.Format;
                key.Append('#').Append(f.FontFamily).Append(',').Append(f.SizeHalfPoints)
                    .Append(f.Bold ? 'B' : 'b').Append(f.Italic ? 'I' : 'i')
                    .Append(f.Underline ? 'U' : 'u').Append(f.Strikeout ? 'S' : 's')
                    .Append(',').Append(f.ForeColor.HasValue ? f.ForeColor.Value.ToArgb() : 0)
                    .Append(',').Append(f.BackColor.HasValue ? f.BackColor.Value.ToArgb() : 0);
            }
            return key.ToString();
        }
        private static Paragraph DiffParagraph(Paragraph oldP, Paragraph newP)
        {
            Paragraph result = CloneParagraph(newP, RevisionKind.None);
            string[] a = Tokens(PlainText(oldP));
            string[] b = Tokens(PlainText(newP));
            List<int[]> ops = Diff(a, b);
            List<Inline> inlines = new List<Inline>();
            RunFormat format = FirstFormat(newP);
            for (int i = 0; i < ops.Count; i++)
            {
                int kind = ops[i][0];
                string text = kind == 1 ? a[ops[i][1]] : b[ops[i][2]];
                TextInline run = new TextInline(text, format.Clone());
                if (kind == 1) DocumentRevisions.MarkDelete(run);
                else if (kind == 2) DocumentRevisions.MarkInsert(run);
                inlines.Add(run);
            }
            result.Inlines.Clear();
            result.Inlines.AddRange(inlines);
            return result;
        }

        private static string[] Tokens(string text)
        {
            List<string> parts = new List<string>();
            int i = 0;
            while (i < text.Length)
            {
                int start = i;
                bool space = char.IsWhiteSpace(text[i]);
                while (i < text.Length && char.IsWhiteSpace(text[i]) == space) i++;
                parts.Add(text.Substring(start, i - start));
            }
            return parts.ToArray();
        }

        /// <summary>0 equal, 1 delete old, 2 insert new.</summary>
        private static List<int[]> Diff(string[] a, string[] b)
        {
            List<int[]> ops = new List<int[]>();
            int[,] lcs;
            if ((long)a.Length * b.Length > 4000000L)
            {
                int n = Math.Min(a.Length, b.Length);
                for (int i = 0; i < n; i++)
                    ops.Add(new int[] { string.Equals(a[i], b[i], StringComparison.Ordinal) ? 0 : 0, i, i });
                for (int i = n; i < a.Length; i++) ops.Add(new int[] { 1, i, -1 });
                for (int i = n; i < b.Length; i++) ops.Add(new int[] { 2, -1, i });
                return ops;
            }
            lcs = new int[a.Length + 1, b.Length + 1];
            for (int i = a.Length - 1; i >= 0; i--)
                for (int j = b.Length - 1; j >= 0; j--)
                    lcs[i, j] = string.Equals(a[i], b[j], StringComparison.Ordinal)
                        ? lcs[i + 1, j + 1] + 1
                        : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);
            int x = 0, y = 0;
            while (x < a.Length && y < b.Length)
            {
                if (string.Equals(a[x], b[y], StringComparison.Ordinal))
                {
                    ops.Add(new int[] { 0, x, y }); x++; y++;
                }
                else if (lcs[x + 1, y] >= lcs[x, y + 1])
                {
                    ops.Add(new int[] { 1, x, -1 }); x++;
                }
                else
                {
                    ops.Add(new int[] { 2, -1, y }); y++;
                }
            }
            while (x < a.Length) { ops.Add(new int[] { 1, x, -1 }); x++; }
            while (y < b.Length) { ops.Add(new int[] { 2, -1, y }); y++; }
            return ops;
        }

        private static Paragraph CloneParagraph(Paragraph source, RevisionKind revision)
        {
            Paragraph p = new Paragraph();
            if (source == null) return p;
            p.Alignment = source.Alignment;
            p.LeftIndentTwips = source.LeftIndentTwips;
            p.RightIndentTwips = source.RightIndentTwips;
            p.FirstLineIndentTwips = source.FirstLineIndentTwips;
            p.SpaceBeforeTwips = source.SpaceBeforeTwips;
            p.SpaceAfterTwips = source.SpaceAfterTwips;
            p.LineSpacing = source.LineSpacing;
            p.Bullet = source.Bullet;
            p.BulletLevel = source.BulletLevel;
            p.Numbered = source.Numbered;
            p.NumberLevel = source.NumberLevel;
            p.PageBreakBefore = source.PageBreakBefore;
            p.SectionIndex = source.SectionIndex;
            p.StyleName = source.StyleName;
            p.RepeatTableHeader = source.RepeatTableHeader;
            p.OutlineLevel = source.OutlineLevel;
            p.NoSpaceBetweenSameStyle = source.NoSpaceBetweenSameStyle;
            p.KeepLinesTogether = source.KeepLinesTogether;
            p.KeepWithNext = source.KeepWithNext;
            p.WidowControl = source.WidowControl;
            p.TabStopsTwips = source.TabStopsTwips == null ? new int[0] : (int[])source.TabStopsTwips.Clone();
            p.TableRows = source.TableRows;
            p.TableWidthTwips = source.TableWidthTwips;
            p.TableStyleName = source.TableStyleName;
            p.TableShading = source.TableShading;
            p.TableBorderColor = source.TableBorderColor;
            p.TableBorderSize = source.TableBorderSize;
            p.TableCells = source.TableCells;
            for (int i = 0; i < source.Inlines.Count; i++)
            {
                TextInline t = source.Inlines[i] as TextInline;
                if (t != null)
                {
                    TextInline copy = t.Clone();
                    if (revision != RevisionKind.None) copy.Revision = revision;
                    p.Inlines.Add(copy);
                }
                else
                {
                    p.Inlines.Add(source.Inlines[i]);
                }
            }
            for (int i = 0; i < source.FloatingObjects.Count; i++)
                p.FloatingObjects.Add(source.FloatingObjects[i].Clone());
            return p;
        }

        private static void CopyDocumentProperties(DocumentModel source, DocumentModel target)
        {
            target.PageWidthTwips = source.PageWidthTwips;
            target.PageHeightTwips = source.PageHeightTwips;
            target.MarginLeftTwips = source.MarginLeftTwips;
            target.MarginRightTwips = source.MarginRightTwips;
            target.MarginTopTwips = source.MarginTopTwips;
            target.MarginBottomTwips = source.MarginBottomTwips;
            target.HeaderText = source.HeaderText;
            target.FooterText = source.FooterText;
            target.Title = source.Title;
            target.Author = source.Author;
            target.HeaderDistanceTwips = source.HeaderDistanceTwips;
            target.FooterDistanceTwips = source.FooterDistanceTwips;
            for (int i = 0; i < source.Sections.Count; i++) target.Sections.Add(source.Sections[i].Clone());
        }

        private static string PlainText(Paragraph p)
        {
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline t = p.Inlines[i] as TextInline;
                if (t != null) text.Append(t.Text);
            }
            return text.ToString();
        }

        private static RunFormat FirstFormat(Paragraph p)
        {
            for (int i = 0; i < p.Inlines.Count; i++)
                if (p.Inlines[i].Format != null) return p.Inlines[i].Format;
            return new RunFormat();
        }
    }
}
