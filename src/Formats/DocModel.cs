using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace WordPadClone.Formats
{
    public enum ParagraphAlignment
    {
        Left = 0,
        Center = 1,
        Right = 2,
        Justify = 3
    }

    /// <summary>
    /// A small, format-neutral document model.  It is intentionally limited to
    /// the constructs that WordPad can display: paragraphs, character runs,
    /// bullets, tab stops and inline pictures.
    /// </summary>
    public sealed class PreservedRelationship
    {
        public string Type = string.Empty;
        public string Target = string.Empty;
        public string TargetMode = string.Empty;
    }

    public sealed class PreservedContentType
    {
        public string Extension = string.Empty;
        public string PartName = string.Empty;
        public string ContentType = string.Empty;
    }

    public sealed class DocumentModel
    {
        public readonly List<Paragraph> Paragraphs = new List<Paragraph>();
        public readonly List<DocumentSection> Sections = new List<DocumentSection>();
        public readonly List<DocumentNote> Footnotes = new List<DocumentNote>();
        public readonly List<DocumentNote> Endnotes = new List<DocumentNote>();
        public readonly List<DocumentComment> Comments = new List<DocumentComment>();
        public readonly List<string> UnsupportedFeatures = new List<string>();
        public readonly List<string> SecurityWarnings = new List<string>();

        public int PageWidthTwips = 12240;   // 8.5 in
        public int PageHeightTwips = 15840;  // 11 in
        public int MarginLeftTwips = 1701;   // 30 mm
        public int MarginRightTwips = 1701;  // 30 mm
        public int MarginTopTwips = 1417;    // 25 mm
        public int MarginBottomTwips = 1417; // 25 mm

        public string Title = string.Empty;
        public string Author = string.Empty;

        // Page furniture is part of the document, not of the application.
        // {PAGE} and {NUMPAGES} are expanded by the view and by the printer.
        public string HeaderText = string.Empty;
        public string FooterText = string.Empty;
        // Package parts that Mini Word does not model (theme, font table,
        // customXml, glossary, ...).  They are copied unchanged when saving so
        // a round-trip does not silently discard the original package data.
        public Dictionary<string, byte[]> PreservedParts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        public List<PreservedRelationship> PreservedRelationships = new List<PreservedRelationship>();
        public List<PreservedContentType> PreservedContentTypes = new List<PreservedContentType>();
        public int HeaderDistanceTwips = 720;
        public int FooterDistanceTwips = 720;

        public bool IsLandscape
        {
            get { return PageWidthTwips > PageHeightTwips; }
        }

        public void EnsureSections()
        {
            if (Sections.Count > 0) return;
            Sections.Add(DocumentSection.FromDocument(this, 0));
        }

        public DocumentSection SectionForParagraph(int paragraphIndex)
        {
            EnsureSections();
            DocumentSection selected = Sections[0];
            for (int i = 0; i < Sections.Count; i++)
            {
                if (Sections[i].StartParagraph <= paragraphIndex) selected = Sections[i];
                else break;
            }
            return selected;
        }

        public void SyncFirstSection()
        {
            EnsureSections();
            DocumentSection s = Sections[0];
            s.PageWidthTwips = PageWidthTwips;
            s.PageHeightTwips = PageHeightTwips;
            s.MarginLeftTwips = MarginLeftTwips;
            s.MarginRightTwips = MarginRightTwips;
            s.MarginTopTwips = MarginTopTwips;
            s.MarginBottomTwips = MarginBottomTwips;
            s.HeaderText = HeaderText;
            s.FooterText = FooterText;
            s.HeaderDistanceTwips = HeaderDistanceTwips;
            s.FooterDistanceTwips = FooterDistanceTwips;
        }

        public string PlainText
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < Paragraphs.Count; i++)
                {
                    if (i > 0) sb.Append("\r\n");
                    Paragraph p = Paragraphs[i];
                    if (p.TableRows != null)
                    {
                        for (int r = 0; r < p.TableRows.Length; r++)
                        {
                            if (r > 0) sb.Append("\r\n");
                            sb.Append(string.Join("\t", p.TableRows[r]));
                        }
                        continue;
                    }
                    for (int j = 0; j < p.Inlines.Count; j++)
                    {
                        TextInline t = p.Inlines[j] as TextInline;
                        if (t != null) sb.Append(t.Text);
                    }
                }
                return sb.ToString();
            }
        }

        public void SetPageSize(int widthTwips, int heightTwips)
        {
            if (widthTwips > 0) PageWidthTwips = widthTwips;
            if (heightTwips > 0) PageHeightTwips = heightTwips;
        }
    }

    public enum SectionBreakType
    {
        NextPage = 0,
        Continuous = 1,
        EvenPage = 2,
        OddPage = 3
    }

    public sealed class DocumentSection
    {
        public int StartParagraph;
        public int PageWidthTwips = 12240;
        public int PageHeightTwips = 15840;
        public int MarginLeftTwips = 1701;
        public int MarginRightTwips = 1701;
        public int MarginTopTwips = 1417;
        public int MarginBottomTwips = 1417;
        public int HeaderDistanceTwips = 720;
        public int FooterDistanceTwips = 720;
        public int ColumnCount = 1;
        public int ColumnSpacingTwips = 720;
        public bool ColumnsSeparator;
        public SectionBreakType BreakType = SectionBreakType.NextPage;
        public string HeaderText = string.Empty;
        public string FooterText = string.Empty;
        public string FirstPageHeaderText = string.Empty;
        public string FirstPageFooterText = string.Empty;
        public string EvenPageHeaderText = string.Empty;
        public string EvenPageFooterText = string.Empty;
        public RunFormat HeaderFormat = new RunFormat("Calibri", 18);
        public RunFormat FooterFormat = new RunFormat("Calibri", 18);
        public RunFormat FirstPageHeaderFormat = new RunFormat("Calibri", 18);
        public RunFormat FirstPageFooterFormat = new RunFormat("Calibri", 18);
        public RunFormat EvenPageHeaderFormat = new RunFormat("Calibri", 18);
        public RunFormat EvenPageFooterFormat = new RunFormat("Calibri", 18);
        public ParagraphAlignment HeaderAlignment = ParagraphAlignment.Left;
        public ParagraphAlignment FooterAlignment = ParagraphAlignment.Left;
        public ParagraphAlignment FirstPageHeaderAlignment = ParagraphAlignment.Left;
        public ParagraphAlignment FirstPageFooterAlignment = ParagraphAlignment.Left;
        public ParagraphAlignment EvenPageHeaderAlignment = ParagraphAlignment.Left;
        public ParagraphAlignment EvenPageFooterAlignment = ParagraphAlignment.Left;
        public List<ImageInline> HeaderImages = new List<ImageInline>();
        public List<ImageInline> FooterImages = new List<ImageInline>();
        public List<ImageInline> FirstPageHeaderImages = new List<ImageInline>();
        public List<ImageInline> FirstPageFooterImages = new List<ImageInline>();
        public List<ImageInline> EvenPageHeaderImages = new List<ImageInline>();
        public List<ImageInline> EvenPageFooterImages = new List<ImageInline>();
        public bool TitlePage;
        public bool DifferentOddEven;

        public static DocumentSection FromDocument(DocumentModel doc, int start)
        {
            DocumentSection s = new DocumentSection();
            s.StartParagraph = start;
            s.PageWidthTwips = doc.PageWidthTwips;
            s.PageHeightTwips = doc.PageHeightTwips;
            s.MarginLeftTwips = doc.MarginLeftTwips;
            s.MarginRightTwips = doc.MarginRightTwips;
            s.MarginTopTwips = doc.MarginTopTwips;
            s.MarginBottomTwips = doc.MarginBottomTwips;
            s.HeaderText = doc.HeaderText;
            s.FooterText = doc.FooterText;
            s.HeaderDistanceTwips = doc.HeaderDistanceTwips;
            s.FooterDistanceTwips = doc.FooterDistanceTwips;
            return s;
        }

        public DocumentSection Clone()
        {
            DocumentSection copy = (DocumentSection)MemberwiseClone();
            copy.HeaderFormat = HeaderFormat == null ? new RunFormat("Calibri", 18) : HeaderFormat.Clone();
            copy.FooterFormat = FooterFormat == null ? new RunFormat("Calibri", 18) : FooterFormat.Clone();
            copy.FirstPageHeaderFormat = FirstPageHeaderFormat == null ? new RunFormat("Calibri", 18) : FirstPageHeaderFormat.Clone();
            copy.FirstPageFooterFormat = FirstPageFooterFormat == null ? new RunFormat("Calibri", 18) : FirstPageFooterFormat.Clone();
            copy.EvenPageHeaderFormat = EvenPageHeaderFormat == null ? new RunFormat("Calibri", 18) : EvenPageHeaderFormat.Clone();
            copy.EvenPageFooterFormat = EvenPageFooterFormat == null ? new RunFormat("Calibri", 18) : EvenPageFooterFormat.Clone();
            copy.HeaderImages = new List<ImageInline>();
            copy.FooterImages = new List<ImageInline>();
            copy.FirstPageHeaderImages = CloneImages(FirstPageHeaderImages);
            copy.FirstPageFooterImages = CloneImages(FirstPageFooterImages);
            copy.EvenPageHeaderImages = CloneImages(EvenPageHeaderImages);
            copy.EvenPageFooterImages = CloneImages(EvenPageFooterImages);
            for (int i = 0; i < HeaderImages.Count; i++)
            {
                ImageInline image = HeaderImages[i];
                copy.HeaderImages.Add(new ImageInline(image.Data, image.Extension, image.WidthTwips,
                    image.HeightTwips, image.Format.Clone()));
            }
            for (int i = 0; i < FooterImages.Count; i++)
            {
                ImageInline image = FooterImages[i];
                copy.FooterImages.Add(new ImageInline(image.Data, image.Extension, image.WidthTwips,
                    image.HeightTwips, image.Format.Clone()));
            }
            return copy;
        }

        private static List<ImageInline> CloneImages(List<ImageInline> source)
        {
            List<ImageInline> result = new List<ImageInline>();
            if (source == null) return result;
            for (int i = 0; i < source.Count; i++)
            {
                ImageInline image = source[i];
                result.Add(new ImageInline(image.Data, image.Extension, image.WidthTwips,
                    image.HeightTwips, image.Format.Clone()));
            }
            return result;
        }
    }

    public sealed class Paragraph
    {
        // Paragraphs with fields/TOC/content controls keep their original
        // WordprocessingML while untouched.  This preserves instruction codes
        // and package references that the editor does not model.
        public string PreservedXml;
        public string PreservedText;
        public string PreservedFormatKey;
        public ParagraphAlignment Alignment = ParagraphAlignment.Left;
        public int LeftIndentTwips;       // indentation of continuation lines
        public int RightIndentTwips;
        public int FirstLineIndentTwips;  // relative to LeftIndentTwips
        public int SpaceBeforeTwips;
        public int SpaceAfterTwips;
        public int LineSpacing;           // 0 = single, 1 = 1.5, 2 = double
        public bool Bullet;
        public int BulletLevel;
        public bool Numbered;
        public int NumberLevel;
        public bool PageBreakBefore;
        public int SectionIndex;
        public int OutlineLevel;
        public bool NoSpaceBetweenSameStyle;
        public bool KeepLinesTogether;
        public bool KeepWithNext;
        public bool WidowControl = true;
        public int[] TabStopsTwips = new int[0];
        public readonly List<Inline> Inlines = new List<Inline>();
        public readonly List<FloatingObject> FloatingObjects = new List<FloatingObject>();

        // When non-null this paragraph is a table block. Each item is one row
        // and each string is one cell. Inlines is empty for table blocks.
        public string[][] TableRows;
        public int TableWidthTwips;
        public string TableStyleName = string.Empty;
        public Color? TableShading;
        public Color? TableBorderColor;
        public int TableBorderSize;

        // Merge-aware cells. TableRows remains the compact representation used
        // by the format readers; TableCells is the editing/rendering model.
        public TableCell[][] TableCells;

        public void EnsureTableCells()
        {
            if (TableCells != null || TableRows == null) return;
            int rows = TableRows.Length;
            int columns = 0;
            for (int r = 0; r < rows; r++)
                if (TableRows[r] != null && TableRows[r].Length > columns) columns = TableRows[r].Length;
            TableCells = new TableCell[rows][];
            for (int r = 0; r < rows; r++)
            {
                TableCells[r] = new TableCell[columns];
                for (int c = 0; c < columns; c++)
                {
                    TableCells[r][c] = new TableCell();
                    TableCells[r][c].Text = TableRows[r] != null && c < TableRows[r].Length ? TableRows[r][c] : string.Empty;
                }
            }
        }

        public void SyncTableRows()
        {
            if (TableCells == null) return;
            int rows = TableCells.Length;
            int columns = 0;
            for (int r = 0; r < rows; r++)
                if (TableCells[r] != null && TableCells[r].Length > columns) columns = TableCells[r].Length;
            TableRows = new string[rows][];
            for (int r = 0; r < rows; r++)
            {
                TableRows[r] = new string[columns];
                for (int c = 0; c < columns; c++)
                {
                    TableCell cell = TableCells[r] != null && c < TableCells[r].Length ? TableCells[r][c] : null;
                    TableRows[r][c] = cell == null || cell.Covered ? string.Empty : (cell.Text ?? string.Empty);
                }
            }
        }
    }

    public sealed class TableCell
    {
        public string Text = string.Empty;
        public int RowSpan = 1;
        public int ColSpan = 1;
        public bool Covered;
        public Color? Shading;
        public Color? BorderColor;
        public int BorderSize;
        public bool Bold;
        public bool Italic;
        public string LinkUrl;
        public RunFormat Format;

        public TableCell Clone()
        {
            return new TableCell
            {
                Text = Text,
                RowSpan = RowSpan,
                ColSpan = ColSpan,
                Covered = Covered,
                Shading = Shading,
                BorderColor = BorderColor,
                BorderSize = BorderSize,
                Bold = Bold,
                Italic = Italic,
                LinkUrl = LinkUrl,
                Format = Format == null ? null : Format.Clone()
            };
        }

    }

    public abstract class Inline
    {
        public RunFormat Format;

        protected Inline(RunFormat format)
        {
            Format = format;
        }
    }

    public sealed class TextInline : Inline
    {
        public string Text;

        public TextInline(string text, RunFormat format)
            : base(format)
        {
            Text = text ?? string.Empty;
        }
    }

    public sealed class ImageInline : Inline
    {
        public byte[] Data;
        public string Extension;   // png, jpeg, wmf, emf, gif, bmp
        public int WidthTwips;
        public int HeightTwips;

        public ImageInline(byte[] data, string extension, int widthTwips, int heightTwips, RunFormat format)
            : base(format)
        {
            Data = data;
            Extension = extension;
            WidthTwips = widthTwips > 0 ? widthTwips : 1440;
            HeightTwips = heightTwips > 0 ? heightTwips : 1440;
        }
    }

    public sealed class DocumentNote
    {
        public int Id;
        public string Text = string.Empty;
        public string Author = string.Empty;
        public bool Endnote;

        public DocumentNote Clone()
        {
            return (DocumentNote)MemberwiseClone();
        }
    }

    public sealed class DocumentComment
    {
        public int Id;
        public string Text = string.Empty;
        public string Author = string.Empty;
        public string Initials = string.Empty;
        public string Date = string.Empty;

        public DocumentComment Clone()
        {
            return (DocumentComment)MemberwiseClone();
        }
    }

    public sealed class NoteInline : Inline
    {
        public DocumentNote Note;
        public NoteInline(DocumentNote note, RunFormat format) : base(format) { Note = note; }
    }

    public sealed class CommentInline : Inline
    {
        public DocumentComment Comment;
        public CommentInline(DocumentComment comment, RunFormat format) : base(format) { Comment = comment; }
    }

    public enum FloatingKind
    {
        Image,
        TextBox,
        Shape,
        Embedded
    }

    public enum FloatingShapeKind
    {
        Rectangle,
        Ellipse,
        Triangle,
        Star,
        Arrow,
        Diamond,
        Heart,
        ArrowLeft
    }

    public enum TextWrapMode
    {
        Inline,
        Square,
        Tight,
        Through,
        TopAndBottom,
        Behind,
        InFront
    }

    public sealed class FloatingObject
    {
        public FloatingKind Kind = FloatingKind.Image;
        public ImageInline Image;
        public string Text = string.Empty;
        public RunFormat Format = new RunFormat("Calibri", 20);
        public Color? Fill;
        public Color? BorderColor;
        public int BorderSize = 1;
        public byte[] EmbeddedData;
        public string EmbeddedName = string.Empty;
        public string ProgId = string.Empty;
        public string ShapeId = string.Empty;
        public FloatingShapeKind ShapeKind = FloatingShapeKind.Rectangle;
        public int WidthTwips = 1440;
        public int HeightTwips = 720;
        public int OffsetXTwips;
        public int OffsetYTwips;
        public TextWrapMode Wrap = TextWrapMode.Square;
        public bool BehindText;
        public int Rotation;

        public FloatingObject Clone()
        {
            FloatingObject copy = (FloatingObject)MemberwiseClone();
            copy.Format = Format == null ? new RunFormat() : Format.Clone();
            if (Image != null)
                copy.Image = new ImageInline(Image.Data, Image.Extension, Image.WidthTwips, Image.HeightTwips, Image.Format.Clone());
            return copy;
        }
    }

    public sealed class RunFormat
    {
        public string FontFamily = "Calibri";
        public int SizeHalfPoints = 22;       // 11 pt
        public bool Bold;
        public bool Italic;
        public bool Underline;
        public bool Strikeout;
        public Color? ForeColor;              // null = automatic (black)
        public Color? BackColor;              // character highlight
        public int Script;                    // -1 subscript, 0 normal, 1 superscript
        public string LinkUrl;

        public RunFormat()
        {
        }

        public RunFormat(string family, int halfPoints)
        {
            if (!string.IsNullOrEmpty(family)) FontFamily = family;
            if (halfPoints > 0) SizeHalfPoints = halfPoints;
        }

        public RunFormat Clone()
        {
            RunFormat f = new RunFormat();
            f.FontFamily = FontFamily;
            f.SizeHalfPoints = SizeHalfPoints;
            f.Bold = Bold;
            f.Italic = Italic;
            f.Underline = Underline;
            f.Strikeout = Strikeout;
            f.ForeColor = ForeColor;
            f.BackColor = BackColor;
            f.Script = Script;
            f.LinkUrl = LinkUrl;
            return f;
        }

        public bool SameAs(RunFormat other)
        {
            if (other == null) return false;
            return string.Equals(FontFamily, other.FontFamily, StringComparison.OrdinalIgnoreCase)
                && SizeHalfPoints == other.SizeHalfPoints
                && Bold == other.Bold
                && Italic == other.Italic
                && Underline == other.Underline
                && Strikeout == other.Strikeout
                && Nullable.Equals(ForeColor, other.ForeColor)
                && Nullable.Equals(BackColor, other.BackColor)
                && Script == other.Script
                && string.Equals(LinkUrl, other.LinkUrl, StringComparison.Ordinal);
        }

        public string Key
        {
            get
            {
                return (FontFamily ?? string.Empty).ToLowerInvariant() + "|" + SizeHalfPoints + "|"
                    + (Bold ? "B" : "") + (Italic ? "I" : "") + (Underline ? "U" : "") + (Strikeout ? "S" : "")
                    + "|" + (ForeColor.HasValue ? ColorTranslator.ToHtml(ForeColor.Value) : "auto")
                    + "|" + (BackColor.HasValue ? ColorTranslator.ToHtml(BackColor.Value) : "none")
                    + "|" + Script;
            }
        }
    }
}
