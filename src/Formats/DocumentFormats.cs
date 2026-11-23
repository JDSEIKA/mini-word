using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using WordPadClone.Util;

namespace WordPadClone.Formats
{
    public enum DocumentFormat
    {
        Rtf,
        Docx,
        Odt,
        Text,
        Doc
    }

    public static class DocumentFormats
    {
        /// <summary>Filter used by Open; .doc is read-only and is converted on save.</summary>
        public const string Filter =
            "Documentos de Word (*.docx)|*.docx|" +
            "Documentos de Word 97-2003 (*.doc;*.dot)|*.doc;*.dot|" +
            "Documentos de OpenDocument (*.odt)|*.odt|" +
            "Documentos de formato RTF (*.rtf)|*.rtf|" +
            "Documentos de texto (*.txt)|*.txt|" +
            "Todos los archivos (*.*)|*.*";

        /// <summary>Filter used by Save As; legacy .doc is intentionally absent.</summary>
        public const string SaveFilter =
            "Documentos de Word (*.docx)|*.docx|" +
            "Documentos de OpenDocument (*.odt)|*.odt|" +
            "Documentos de formato RTF (*.rtf)|*.rtf|" +
            "Documentos de texto (*.txt)|*.txt";

        public static string DefaultExtension(DocumentFormat format)
        {
            switch (format)
            {
                case DocumentFormat.Docx: return ".docx";
                case DocumentFormat.Odt: return ".odt";
                case DocumentFormat.Text: return ".txt";
                case DocumentFormat.Doc: return ".docx";
                default: return ".rtf";
            }
        }

        public static DocumentFormat FromExtension(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            switch (ext)
            {
                case ".docx": return DocumentFormat.Docx;
                case ".doc":
                case ".dot":
                    return DocumentFormat.Doc;
                case ".odt": return DocumentFormat.Odt;
                case ".txt":
                case ".log":
                case ".csv":
                    return DocumentFormat.Text;
                default:
                    return DocumentFormat.Rtf;
            }
        }

        public static DocumentFormat Detect(string path)
        {
            try
            {
                byte[] head = ReadHead(path, 512);
                if (head.Length >= 8 && head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0
                    && head[4] == 0xA1 && head[5] == 0xB1 && head[6] == 0x1A && head[7] == 0xE1)
                    return DocumentFormat.Doc;
                if (head.Length >= 4 && head[0] == 0x50 && head[1] == 0x4B && head[2] == 0x03 && head[3] == 0x04)
                {
                    Dictionary<string, byte[]> zip = ZipArchiveLite.Read(path);
                    if (zip.ContainsKey("word/document.xml")) return DocumentFormat.Docx;
                    if (zip.ContainsKey("content.xml")) return DocumentFormat.Odt;
                }
                int i = 0;
                if (head.Length >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF) i = 3;
                while (i < head.Length && (head[i] == ' ' || head[i] == '\r' || head[i] == '\n' || head[i] == '\t')) i++;
                if (i + 4 < head.Length && head[i] == '{' && head[i + 1] == '\\' && head[i + 2] == 'r' && head[i + 3] == 't' && head[i + 4] == 'f')
                    return DocumentFormat.Rtf;
            }
            catch
            {
            }
            return FromExtension(path);
        }

        private static byte[] ReadHead(string path, int max)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                int n = (int)Math.Min(max, fs.Length);
                byte[] buffer = new byte[n];
                int read = 0;
                while (read < n)
                {
                    int r = fs.Read(buffer, read, n - read);
                    if (r <= 0) break;
                    read += r;
                }
                if (read == n) return buffer;
                byte[] small = new byte[read];
                Buffer.BlockCopy(buffer, 0, small, 0, read);
                return small;
            }
        }

        public static string ReadRtfText(string path)
        {
            FileInfo info = new FileInfo(path);
            if (info.Length > 64L * 1024 * 1024)
                throw new InvalidDataException("El RTF supera el límite de 64 MB.");
            byte[] bytes = File.ReadAllBytes(path);
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return RtfSanitizer.Normalize(new UTF8Encoding(false).GetString(bytes, 3, bytes.Length - 3));
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return RtfSanitizer.Normalize(Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2));
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return RtfSanitizer.Normalize(Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2));
            Encoding enc = Encoding.GetEncoding(1252);
            return RtfSanitizer.Normalize(enc.GetString(bytes));
        }

        public static DocumentModel ReadModel(string path, DocumentFormat format)
        {
            FileInfo info = new FileInfo(path);
            if (info.Length > 64L * 1024 * 1024)
                throw new InvalidDataException("El documento supera el límite de 64 MB.");
            switch (format)
            {
                case DocumentFormat.Docx:
                    return DocxFormat.Read(path);
                case DocumentFormat.Odt:
                    return OdtFormat.Read(path);
                case DocumentFormat.Rtf:
                    return RtfReader.Read(ReadRtfText(path));
                case DocumentFormat.Doc:
                    return DocFormat.Read(path);
                case DocumentFormat.Text:
                    return FromPlainText(TextFile.Load(path));
                default:
                    throw new InvalidOperationException("Formato no compatible.");
            }
        }

        public static DocumentModel FromPlainText(string text)
        {
            DocumentModel doc = new DocumentModel();
            text = TextFile.NormalizeNewlines(text);
            string[] lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                Paragraph p = new Paragraph();
                p.Inlines.Add(new TextInline(lines[i], new RunFormat()));
                doc.Paragraphs.Add(p);
            }
            if (doc.Paragraphs.Count == 0) doc.Paragraphs.Add(new Paragraph());
            return doc;
        }

        public static void SaveDocx(string path, DocumentModel model)
        {
            AtomicFile.Save(path, delegate(string temporary)
            {
                DocxFormat.Write(temporary, NormalizeImages(model));
            });
        }

        public static void SaveOdt(string path, DocumentModel model)
        {
            AtomicFile.Save(path, delegate(string temporary)
            {
                OdtFormat.Write(temporary, NormalizeImages(model));
            });
        }

        public static void SaveRtf(string path, string rtf)
        {
            AtomicFile.Save(path, delegate(string temporary)
            {
                File.WriteAllText(temporary, rtf, new UTF8Encoding(false));
            });
        }

        public static void SaveText(string path, string text, TextEncodingKind encoding)
        {
            AtomicFile.Save(path, delegate(string temporary)
            {
                TextFile.Save(temporary, text, encoding);
            });
        }

        public static DocumentModel NormalizeImages(DocumentModel source)
        {
            DocumentModel copy = new DocumentModel();
            copy.PageWidthTwips = source.PageWidthTwips;
            copy.PageHeightTwips = source.PageHeightTwips;
            copy.MarginLeftTwips = source.MarginLeftTwips;
            copy.MarginRightTwips = source.MarginRightTwips;
            copy.MarginTopTwips = source.MarginTopTwips;
            copy.MarginBottomTwips = source.MarginBottomTwips;
            copy.Title = source.Title;
            copy.Author = source.Author;
            copy.TrackChanges = source.TrackChanges;
            copy.ShowRevisions = source.ShowRevisions;
            copy.HeaderText = source.HeaderText;
            copy.FooterText = source.FooterText;
            copy.HeaderDistanceTwips = source.HeaderDistanceTwips;
            copy.FooterDistanceTwips = source.FooterDistanceTwips;
            source.EnsureSections();
            for (int i = 0; i < source.Sections.Count; i++) copy.Sections.Add(source.Sections[i].Clone());
            for (int i = 0; i < source.Footnotes.Count; i++) copy.Footnotes.Add(source.Footnotes[i].Clone());
            for (int i = 0; i < source.Endnotes.Count; i++) copy.Endnotes.Add(source.Endnotes[i].Clone());
            for (int i = 0; i < source.Comments.Count; i++) copy.Comments.Add(source.Comments[i].Clone());
            for (int i = 0; i < source.UnsupportedFeatures.Count; i++)
                copy.UnsupportedFeatures.Add(source.UnsupportedFeatures[i]);
            for (int i = 0; i < source.SecurityWarnings.Count; i++)
                copy.SecurityWarnings.Add(source.SecurityWarnings[i]);
            if (source.PreservedParts != null)
            {
                foreach (KeyValuePair<string, byte[]> part in source.PreservedParts)
                    copy.PreservedParts[part.Key] = part.Value;
            }
            if (source.PreservedRelationships != null)
            {
                for (int i = 0; i < source.PreservedRelationships.Count; i++)
                {
                    PreservedRelationship original = source.PreservedRelationships[i];
                    PreservedRelationship item = new PreservedRelationship();
                    item.Type = original.Type;
                    item.Target = original.Target;
                    item.TargetMode = original.TargetMode;
                    copy.PreservedRelationships.Add(item);
                }
            }
            if (source.PreservedContentTypes != null)
            {
                for (int i = 0; i < source.PreservedContentTypes.Count; i++)
                {
                    PreservedContentType original = source.PreservedContentTypes[i];
                    PreservedContentType item = new PreservedContentType();
                    item.Extension = original.Extension;
                    item.PartName = original.PartName;
                    item.ContentType = original.ContentType;
                    copy.PreservedContentTypes.Add(item);
                }
            }

            for (int i = 0; i < source.Paragraphs.Count; i++)
            {
                Paragraph sp = source.Paragraphs[i];
                Paragraph p = new Paragraph();
                p.PreservedXml = sp.PreservedXml;
                p.PreservedText = sp.PreservedText;
			p.PreservedFormatKey = sp.PreservedFormatKey;
                p.HasModeledField = sp.HasModeledField;
                p.IsTocEntry = sp.IsTocEntry;
                p.IsTocTitle = sp.IsTocTitle;
                p.TocLevel = sp.TocLevel;
                p.TocTargetParagraph = sp.TocTargetParagraph;
                p.TocPage = sp.TocPage;
                p.Alignment = sp.Alignment;
                p.LeftIndentTwips = sp.LeftIndentTwips;
                p.RightIndentTwips = sp.RightIndentTwips;
                p.FirstLineIndentTwips = sp.FirstLineIndentTwips;
                p.SpaceBeforeTwips = sp.SpaceBeforeTwips;
                p.SpaceAfterTwips = sp.SpaceAfterTwips;
                p.LineSpacing = sp.LineSpacing;
                p.Bullet = sp.Bullet;
                p.BulletLevel = sp.BulletLevel;
                p.Numbered = sp.Numbered;
                p.NumberLevel = sp.NumberLevel;
                p.PageBreakBefore = sp.PageBreakBefore;
                p.SectionIndex = sp.SectionIndex;
                p.OutlineLevel = sp.OutlineLevel;
                p.NoSpaceBetweenSameStyle = sp.NoSpaceBetweenSameStyle;
                p.KeepLinesTogether = sp.KeepLinesTogether;
                p.KeepWithNext = sp.KeepWithNext;
                p.WidowControl = sp.WidowControl;
                p.TabStopsTwips = (int[])sp.TabStopsTwips.Clone();
                p.TableRows = sp.TableRows;
                p.TableWidthTwips = sp.TableWidthTwips;
                p.TableStyleName = sp.TableStyleName;
                p.TableShading = sp.TableShading;
                p.TableBorderColor = sp.TableBorderColor;
                p.TableBorderSize = sp.TableBorderSize;
                if (sp.TableCells != null)
                {
                    p.TableCells = new TableCell[sp.TableCells.Length][];
                    for (int r = 0; r < sp.TableCells.Length; r++)
                    {
                        p.TableCells[r] = sp.TableCells[r] == null ? new TableCell[0] : new TableCell[sp.TableCells[r].Length];
                        for (int c = 0; c < p.TableCells[r].Length; c++)
                            p.TableCells[r][c] = sp.TableCells[r][c] == null ? null : sp.TableCells[r][c].Clone();
                    }
                }

                for (int j = 0; j < sp.Inlines.Count; j++)
                {
                    TextInline t = sp.Inlines[j] as TextInline;
                    if (t != null)
                    {
                        p.Inlines.Add(t.Clone());
                        continue;
                    }
                    ImageInline img = sp.Inlines[j] as ImageInline;
                    if (img != null)
                    {
                        ImageInline normalized = ImageUtil.NormalizeForPackage(img);
                        normalized.Format = img.Format.Clone();
                        p.Inlines.Add(normalized);
                        continue;
                    }
                    NoteInline note = sp.Inlines[j] as NoteInline;
                    if (note != null) { p.Inlines.Add(new NoteInline(note.Note, note.Format.Clone())); continue; }
                    CommentInline comment = sp.Inlines[j] as CommentInline;
                    if (comment != null) p.Inlines.Add(new CommentInline(comment.Comment, comment.Format.Clone()));
                }
                for (int j = 0; j < sp.FloatingObjects.Count; j++)
                {
                    FloatingObject floating = sp.FloatingObjects[j].Clone();
                    if (floating.Image != null)
                    {
                        ImageInline normalized = ImageUtil.NormalizeForPackage(floating.Image);
                        normalized.Format = floating.Image.Format.Clone();
                        floating.Image = normalized;
                    }
                    p.FloatingObjects.Add(floating);
                }
                copy.Paragraphs.Add(p);
            }
            if (copy.Paragraphs.Count == 0) copy.Paragraphs.Add(new Paragraph());
            return copy;
        }
    }
}
