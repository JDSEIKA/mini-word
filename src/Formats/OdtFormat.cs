using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace WordPadClone.Formats
{
    /// <summary>OpenDocument Text reader and writer (flat package, ODF 1.2).</summary>
    public static class OdtFormat
    {
        private static readonly XNamespace OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        private static readonly XNamespace STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        private static readonly XNamespace TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private static readonly XNamespace TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        private static readonly XNamespace DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        private static readonly XNamespace FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        private static readonly XNamespace XLINK = "http://www.w3.org/1999/xlink";
        private static readonly XNamespace SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        private static readonly XNamespace META = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
        private static readonly XNamespace DC = "http://purl.org/dc/elements/1.1/";
        private static readonly XNamespace MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        private static readonly XNamespace CONFIG = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

        // ------------------------------------------------------------------
        // Writing
        // ------------------------------------------------------------------

        private sealed class PStyle
        {
            public int Id;
            public ParagraphAlignment Alignment;
            public int Left;
            public int Right;
            public int First;
            public int Before;
            public int After;
            public int LineSpacing;
            public bool Bullet;
        }

        private sealed class TStyle
        {
            public int Id;
            public RunFormat Format;
            public string Link;
        }

        private sealed class ImageInfo
        {
            public string FileName;
            public string Extension;
            public byte[] Data;
            public int WidthTwips;
            public int HeightTwips;
        }

        public static void Write(string path, DocumentModel doc)
        {
            List<PStyle> pStyles = new List<PStyle>();
            List<TStyle> tStyles = new List<TStyle>();
            List<ImageInfo> images = new List<ImageInfo>();
            Dictionary<string, int> pKeys = new Dictionary<string, int>();
            Dictionary<string, int> tKeys = new Dictionary<string, int>();

            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Paragraph p = doc.Paragraphs[i];
                string pkey = ParagraphKey(p);
                if (!pKeys.ContainsKey(pkey))
                {
                    PStyle ps = new PStyle();
                    ps.Id = pStyles.Count + 1;
                    ps.Alignment = p.Alignment;
                    ps.Bullet = p.Bullet;
                    if (!p.Bullet)
                    {
                        ps.Left = p.LeftIndentTwips;
                        ps.Right = p.RightIndentTwips;
                        ps.First = p.FirstLineIndentTwips;
                    }
                    ps.Before = p.SpaceBeforeTwips;
                    ps.After = p.SpaceAfterTwips;
                    ps.LineSpacing = p.LineSpacing;
                    pStyles.Add(ps);
                    pKeys[pkey] = pStyles.Count - 1;
                }

                for (int j = 0; j < p.Inlines.Count; j++)
                {
                    TextInline t = p.Inlines[j] as TextInline;
                    if (t != null)
                    {
                        string tkey = t.Format.Key + "|" + (t.Format.LinkUrl ?? string.Empty);
                        if (!tKeys.ContainsKey(tkey))
                        {
                            TStyle ts = new TStyle();
                            ts.Id = tStyles.Count + 1;
                            ts.Format = t.Format.Clone();
                            ts.Link = t.Format.LinkUrl;
                            tStyles.Add(ts);
                            tKeys[tkey] = tStyles.Count - 1;
                        }
                    }
                    else
                    {
                        ImageInline img = p.Inlines[j] as ImageInline;
                        if (img != null && img.Data != null && img.Data.Length > 0)
                        {
                            ImageInline normalized = ImageUtil.NormalizeForPackage(img);
                            ImageInfo info = FindImage(images, normalized.Data);
                            if (info == null)
                            {
                                info = new ImageInfo();
                                info.Extension = NormalizeExt(normalized.Extension);
                                info.Data = normalized.Data;
                                info.WidthTwips = normalized.WidthTwips;
                                info.HeightTwips = normalized.HeightTwips;
                                info.FileName = "image" + (images.Count + 1) + "." + info.Extension;
                                images.Add(info);
                            }
                        }
                    }
                }
            }

            List<ZipEntryData> entries = new List<ZipEntryData>();
            entries.Add(new ZipEntryData("mimetype", "application/vnd.oasis.opendocument.text", false));
            entries.Add(new ZipEntryData("content.xml", BuildContentXml(doc, pStyles, pKeys, tStyles, tKeys, images), true));
            entries.Add(new ZipEntryData("styles.xml", BuildStylesXml(doc, tStyles), true));
            entries.Add(new ZipEntryData("meta.xml", BuildMetaXml(doc), true));
            entries.Add(new ZipEntryData("settings.xml", BuildSettingsXml(), true));
            entries.Add(new ZipEntryData("META-INF/manifest.xml", BuildManifest(images), true));
            for (int i = 0; i < images.Count; i++)
                entries.Add(new ZipEntryData("Pictures/" + images[i].FileName, images[i].Data, true));

            ZipArchiveLite.Write(path, entries);
        }

        private static string ParagraphKey(Paragraph p)
        {
            return p.Bullet + "|" + p.Alignment + "|" + p.LeftIndentTwips + "|" + p.RightIndentTwips + "|"
                + p.FirstLineIndentTwips + "|" + p.SpaceBeforeTwips + "|" + p.SpaceAfterTwips + "|" + p.LineSpacing;
        }

        private static ImageInfo FindImage(List<ImageInfo> images, byte[] data)
        {
            for (int i = 0; i < images.Count; i++)
            {
                if (images[i].Data.Length != data.Length) continue;
                bool same = true;
                for (int j = 0; j < data.Length; j++)
                    if (images[i].Data[j] != data[j]) { same = false; break; }
                if (same) return images[i];
            }
            return null;
        }

        private static string NormalizeExt(string ext)
        {
            ext = (ext ?? "png").ToLowerInvariant();
            if (ext == "jpg") ext = "jpeg";
            if (ext != "png" && ext != "jpeg" && ext != "gif" && ext != "bmp") ext = "png";
            return ext;
        }

        private static string MediaTypeFor(string ext)
        {
            switch (ext)
            {
                case "jpeg": return "image/jpeg";
                case "gif": return "image/gif";
                case "bmp": return "image/bmp";
                default: return "image/png";
            }
        }

        private static byte[] BuildContentXml(DocumentModel doc, List<PStyle> pStyles, Dictionary<string, int> pKeys,
            List<TStyle> tStyles, Dictionary<string, int> tKeys, List<ImageInfo> images)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("office", "document-content", OFFICE.NamespaceName);
                    WriteNamespaces(w);
                    w.WriteAttributeString("office", "version", OFFICE.NamespaceName, "1.2");

                    w.WriteStartElement("office", "font-face-decls", OFFICE.NamespaceName);
                    List<string> fonts = new List<string>();
                    fonts.Add("Calibri");
                    for (int i = 0; i < tStyles.Count; i++)
                    {
                        string f = tStyles[i].Format.FontFamily;
                        if (!string.IsNullOrEmpty(f) && !fonts.Contains(f)) fonts.Add(f);
                    }
                    for (int i = 0; i < fonts.Count; i++)
                    {
                        w.WriteStartElement("style", "font-face", STYLE.NamespaceName);
                        w.WriteAttributeString("style", "name", STYLE.NamespaceName, fonts[i]);
                        w.WriteAttributeString("svg", "font-family", SVG.NamespaceName, "'" + fonts[i] + "'");
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();

                    w.WriteStartElement("office", "automatic-styles", OFFICE.NamespaceName);
                    for (int i = 0; i < pStyles.Count; i++) WriteParagraphStyle(w, pStyles[i]);
                    for (int i = 0; i < tStyles.Count; i++) WriteTextStyle(w, tStyles[i]);
                    WriteGraphicStyle(w);
                    w.WriteEndElement();

                    w.WriteStartElement("office", "body", OFFICE.NamespaceName);
                    w.WriteStartElement("office", "text", OFFICE.NamespaceName);

                    int index = 0;
                    while (index < doc.Paragraphs.Count)
                    {
                        Paragraph p = doc.Paragraphs[index];
                        if (!p.Bullet)
                        {
                            WriteParagraph(w, p, pStyles[pKeys[ParagraphKey(p)]].Id,
                                tStyles, tKeys, images);
                            index++;
                            continue;
                        }

                        int level = Math.Max(0, Math.Min(2, p.BulletLevel));
                        w.WriteStartElement("text", "list", TEXT.NamespaceName);
                        w.WriteAttributeString("text", "style-name", TEXT.NamespaceName, "L" + (level + 1));
                        while (index < doc.Paragraphs.Count)
                        {
                            Paragraph item = doc.Paragraphs[index];
                            if (!item.Bullet || Math.Max(0, Math.Min(2, item.BulletLevel)) != level) break;
                            w.WriteStartElement("text", "list-item", TEXT.NamespaceName);
                            WriteParagraph(w, item, pStyles[pKeys[ParagraphKey(item)]].Id, tStyles, tKeys, images);
                            w.WriteEndElement();
                            index++;
                        }
                        w.WriteEndElement();
                    }

                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static void WriteNamespaces(XmlWriter w)
        {
            w.WriteAttributeString("xmlns", "style", null, STYLE.NamespaceName);
            w.WriteAttributeString("xmlns", "text", null, TEXT.NamespaceName);
            w.WriteAttributeString("xmlns", "table", null, TABLE.NamespaceName);
            w.WriteAttributeString("xmlns", "draw", null, DRAW.NamespaceName);
            w.WriteAttributeString("xmlns", "fo", null, FO.NamespaceName);
            w.WriteAttributeString("xmlns", "xlink", null, XLINK.NamespaceName);
            w.WriteAttributeString("xmlns", "svg", null, SVG.NamespaceName);
            w.WriteAttributeString("xmlns", "meta", null, META.NamespaceName);
            w.WriteAttributeString("xmlns", "dc", null, DC.NamespaceName);
            w.WriteAttributeString("xmlns", "config", null, CONFIG.NamespaceName);
        }

        private static void WriteParagraphStyle(XmlWriter w, PStyle s)
        {
            w.WriteStartElement("style", "style", STYLE.NamespaceName);
            w.WriteAttributeString("style", "name", STYLE.NamespaceName, "P" + s.Id);
            w.WriteAttributeString("style", "family", STYLE.NamespaceName, "paragraph");
            w.WriteAttributeString("style", "parent-style-name", STYLE.NamespaceName, "Standard");

            w.WriteStartElement("style", "paragraph-properties", STYLE.NamespaceName);
            if (s.Alignment == ParagraphAlignment.Center)
                w.WriteAttributeString("fo", "text-align", FO.NamespaceName, "center");
            else if (s.Alignment == ParagraphAlignment.Right)
                w.WriteAttributeString("fo", "text-align", FO.NamespaceName, "end");
            else if (s.Alignment == ParagraphAlignment.Justify)
                w.WriteAttributeString("fo", "text-align", FO.NamespaceName, "justify");
            if (!s.Bullet)
            {
                if (s.Left != 0) w.WriteAttributeString("fo", "margin-left", FO.NamespaceName, Inches(s.Left));
                if (s.Right != 0) w.WriteAttributeString("fo", "margin-right", FO.NamespaceName, Inches(s.Right));
                if (s.First != 0) w.WriteAttributeString("fo", "text-indent", FO.NamespaceName, Inches(s.First));
            }
            if (s.Before != 0) w.WriteAttributeString("fo", "margin-top", FO.NamespaceName, Inches(s.Before));
            if (s.After != 0) w.WriteAttributeString("fo", "margin-bottom", FO.NamespaceName, Inches(s.After));
            if (s.LineSpacing == 1) w.WriteAttributeString("fo", "line-height", FO.NamespaceName, "150%");
            else if (s.LineSpacing == 2) w.WriteAttributeString("fo", "line-height", FO.NamespaceName, "200%");
            w.WriteEndElement();

            w.WriteEndElement();
        }

        private static void WriteTextStyle(XmlWriter w, TStyle s)
        {
            RunFormat f = s.Format;
            w.WriteStartElement("style", "style", STYLE.NamespaceName);
            w.WriteAttributeString("style", "name", STYLE.NamespaceName, "T" + s.Id);
            w.WriteAttributeString("style", "family", STYLE.NamespaceName, "text");
            w.WriteStartElement("style", "text-properties", STYLE.NamespaceName);
            if (!string.IsNullOrEmpty(f.FontFamily))
                w.WriteAttributeString("style", "font-name", STYLE.NamespaceName, f.FontFamily);
            w.WriteAttributeString("fo", "font-size", FO.NamespaceName, Points(f.SizeHalfPoints));
            if (f.Bold) w.WriteAttributeString("fo", "font-weight", FO.NamespaceName, "bold");
            if (f.Italic) w.WriteAttributeString("fo", "font-style", FO.NamespaceName, "italic");
            if (f.Underline) w.WriteAttributeString("style", "text-underline-style", STYLE.NamespaceName, "solid");
            if (f.Strikeout) w.WriteAttributeString("style", "text-line-through-style", STYLE.NamespaceName, "solid");
            if (f.ForeColor.HasValue) w.WriteAttributeString("fo", "color", FO.NamespaceName, Hex(f.ForeColor.Value));
            if (f.BackColor.HasValue) w.WriteAttributeString("fo", "background-color", FO.NamespaceName, Hex(f.BackColor.Value));
            if (f.Script > 0) w.WriteAttributeString("style", "text-position", STYLE.NamespaceName, "super 58%");
            else if (f.Script < 0) w.WriteAttributeString("style", "text-position", STYLE.NamespaceName, "sub 58%");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteGraphicStyle(XmlWriter w)
        {
            w.WriteStartElement("style", "style", STYLE.NamespaceName);
            w.WriteAttributeString("style", "name", STYLE.NamespaceName, "G1");
            w.WriteAttributeString("style", "family", STYLE.NamespaceName, "graphic");
            w.WriteStartElement("style", "graphic-properties", STYLE.NamespaceName);
            w.WriteAttributeString("style", "run-through", STYLE.NamespaceName, "foreground");
            w.WriteAttributeString("style", "wrap", STYLE.NamespaceName, "none");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteParagraph(XmlWriter w, Paragraph p, int styleId, List<TStyle> tStyles,
            Dictionary<string, int> tKeys, List<ImageInfo> images)
        {
            w.WriteStartElement("text", "p", TEXT.NamespaceName);
            w.WriteAttributeString("text", "style-name", TEXT.NamespaceName, "P" + styleId);

            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline t = p.Inlines[i] as TextInline;
                if (t != null)
                {
                    string key = t.Format.Key + "|" + (t.Format.LinkUrl ?? string.Empty);
                    int tid = tStyles[tKeys[key]].Id;
                    bool link = !string.IsNullOrEmpty(t.Format.LinkUrl);
                    if (link)
                    {
                        w.WriteStartElement("text", "a", TEXT.NamespaceName);
                        w.WriteAttributeString("xlink", "type", XLINK.NamespaceName, "simple");
                        w.WriteAttributeString("xlink", "href", XLINK.NamespaceName, t.Format.LinkUrl);
                    }
                    w.WriteStartElement("text", "span", TEXT.NamespaceName);
                    w.WriteAttributeString("text", "style-name", TEXT.NamespaceName, "T" + tid);
                    WriteTextContent(w, t.Text);
                    w.WriteEndElement();
                    if (link) w.WriteEndElement();
                    continue;
                }

                ImageInline img = p.Inlines[i] as ImageInline;
                if (img != null && img.Data != null && img.Data.Length > 0)
                {
                    ImageInline normalized = ImageUtil.NormalizeForPackage(img);
                    ImageInfo info = FindImage(images, normalized.Data);
                    if (info == null) continue;
                    WriteImage(w, info);
                }
            }

            w.WriteEndElement();
        }

        private static void WriteTextContent(XmlWriter w, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            int i = 0;
            while (i < text.Length)
            {
                char c = text[i];
                if (c == '\n')
                {
                    w.WriteStartElement("text", "line-break", TEXT.NamespaceName);
                    w.WriteEndElement();
                    i++;
                }
                else if (c == '\t')
                {
                    w.WriteStartElement("text", "tab", TEXT.NamespaceName);
                    w.WriteEndElement();
                    i++;
                }
                else if (c == ' ')
                {
                    int start = i;
                    while (i < text.Length && text[i] == ' ') i++;
                    int count = i - start;
                    if (count > 1 || start == 0 || i == text.Length)
                    {
                        w.WriteStartElement("text", "s", TEXT.NamespaceName);
                        if (count > 1) w.WriteAttributeString("text", "c", TEXT.NamespaceName, count.ToString(CultureInfo.InvariantCulture));
                        w.WriteEndElement();
                    }
                    else
                    {
                        w.WriteString(" ");
                    }
                }
                else
                {
                    int start = i;
                    while (i < text.Length && text[i] != ' ' && text[i] != '\n' && text[i] != '\t') i++;
                    w.WriteString(text.Substring(start, i - start));
                }
            }
        }

        private static void WriteImage(XmlWriter w, ImageInfo info)
        {
            w.WriteStartElement("draw", "frame", DRAW.NamespaceName);
            w.WriteAttributeString("draw", "style-name", DRAW.NamespaceName, "G1");
            w.WriteAttributeString("draw", "name", DRAW.NamespaceName, "Imagen " + info.FileName);
            w.WriteAttributeString("text", "anchor-type", TEXT.NamespaceName, "as-char");
            w.WriteAttributeString("svg", "width", SVG.NamespaceName, Inches(info.WidthTwips));
            w.WriteAttributeString("svg", "height", SVG.NamespaceName, Inches(info.HeightTwips));
            w.WriteStartElement("draw", "image", DRAW.NamespaceName);
            w.WriteAttributeString("xlink", "href", XLINK.NamespaceName, "Pictures/" + info.FileName);
            w.WriteAttributeString("xlink", "type", XLINK.NamespaceName, "simple");
            w.WriteAttributeString("xlink", "show", XLINK.NamespaceName, "embed");
            w.WriteAttributeString("xlink", "actuate", XLINK.NamespaceName, "onLoad");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static byte[] BuildStylesXml(DocumentModel doc, List<TStyle> tStyles)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("office", "document-styles", OFFICE.NamespaceName);
                    WriteNamespaces(w);
                    w.WriteAttributeString("office", "version", OFFICE.NamespaceName, "1.2");

                    w.WriteStartElement("office", "font-face-decls", OFFICE.NamespaceName);
                    List<string> fonts = new List<string>();
                    fonts.Add("Calibri");
                    for (int i = 0; i < tStyles.Count; i++)
                    {
                        string f = tStyles[i].Format.FontFamily;
                        if (!string.IsNullOrEmpty(f) && !fonts.Contains(f)) fonts.Add(f);
                    }
                    for (int i = 0; i < fonts.Count; i++)
                    {
                        w.WriteStartElement("style", "font-face", STYLE.NamespaceName);
                        w.WriteAttributeString("style", "name", STYLE.NamespaceName, fonts[i]);
                        w.WriteAttributeString("svg", "font-family", SVG.NamespaceName, "'" + fonts[i] + "'");
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();

                    w.WriteStartElement("office", "styles", OFFICE.NamespaceName);
                    w.WriteStartElement("style", "default-style", STYLE.NamespaceName);
                    w.WriteAttributeString("style", "family", STYLE.NamespaceName, "paragraph");
                    w.WriteStartElement("style", "paragraph-properties", STYLE.NamespaceName);
                    w.WriteAttributeString("fo", "orphans", FO.NamespaceName, "2");
                    w.WriteAttributeString("fo", "widows", FO.NamespaceName, "2");
                    w.WriteEndElement();
                    w.WriteStartElement("style", "text-properties", STYLE.NamespaceName);
                    w.WriteAttributeString("style", "font-name", STYLE.NamespaceName, "Calibri");
                    w.WriteAttributeString("fo", "font-size", FO.NamespaceName, "11pt");
                    w.WriteEndElement();
                    w.WriteEndElement();

                    w.WriteStartElement("style", "style", STYLE.NamespaceName);
                    w.WriteAttributeString("style", "name", STYLE.NamespaceName, "Standard");
                    w.WriteAttributeString("style", "family", STYLE.NamespaceName, "paragraph");
                    w.WriteAttributeString("style", "class", STYLE.NamespaceName, "text");
                    w.WriteEndElement();

                    for (int level = 1; level <= 3; level++)
                    {
                        w.WriteStartElement("text", "list-style", TEXT.NamespaceName);
                        w.WriteAttributeString("style", "name", STYLE.NamespaceName, "L" + level);
                        w.WriteStartElement("text", "list-level-style-bullet", TEXT.NamespaceName);
                        w.WriteAttributeString("text", "level", TEXT.NamespaceName, "1");
                        w.WriteAttributeString("text", "bullet-char", TEXT.NamespaceName, "\u2022");
                        w.WriteStartElement("style", "list-level-properties", STYLE.NamespaceName);
                        w.WriteAttributeString("text", "space-before", TEXT.NamespaceName, Inches(720 * (level - 1) + 360));
                        w.WriteAttributeString("text", "min-label-width", TEXT.NamespaceName, "0.25in");
                        w.WriteEndElement();
                        w.WriteStartElement("style", "text-properties", STYLE.NamespaceName);
                        w.WriteAttributeString("style", "font-name", STYLE.NamespaceName, "Calibri");
                        w.WriteEndElement();
                        w.WriteEndElement();
                        w.WriteEndElement();
                    }

                    w.WriteEndElement();

                    w.WriteStartElement("office", "automatic-styles", OFFICE.NamespaceName);
                    w.WriteStartElement("style", "page-layout", STYLE.NamespaceName);
                    w.WriteAttributeString("style", "name", STYLE.NamespaceName, "pm1");
                    w.WriteStartElement("style", "page-layout-properties", STYLE.NamespaceName);
                    bool landscape = doc.IsLandscape;
                    int pw = landscape ? Math.Max(doc.PageWidthTwips, doc.PageHeightTwips) : doc.PageWidthTwips;
                    int ph = landscape ? Math.Min(doc.PageWidthTwips, doc.PageHeightTwips) : doc.PageHeightTwips;
                    w.WriteAttributeString("fo", "page-width", FO.NamespaceName, Inches(pw));
                    w.WriteAttributeString("fo", "page-height", FO.NamespaceName, Inches(ph));
                    w.WriteAttributeString("style", "print-orientation", STYLE.NamespaceName, landscape ? "landscape" : "portrait");
                    w.WriteAttributeString("fo", "margin-top", FO.NamespaceName, Inches(doc.MarginTopTwips));
                    w.WriteAttributeString("fo", "margin-bottom", FO.NamespaceName, Inches(doc.MarginBottomTwips));
                    w.WriteAttributeString("fo", "margin-left", FO.NamespaceName, Inches(doc.MarginLeftTwips));
                    w.WriteAttributeString("fo", "margin-right", FO.NamespaceName, Inches(doc.MarginRightTwips));
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();

                    w.WriteStartElement("office", "master-styles", OFFICE.NamespaceName);
                    w.WriteStartElement("style", "master-page", STYLE.NamespaceName);
                    w.WriteAttributeString("style", "name", STYLE.NamespaceName, "Standard");
                    w.WriteAttributeString("style", "page-layout-name", STYLE.NamespaceName, "pm1");
                    w.WriteEndElement();
                    w.WriteEndElement();

                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildMetaXml(DocumentModel doc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("office", "document-meta", OFFICE.NamespaceName);
                    WriteNamespaces(w);
                    w.WriteAttributeString("office", "version", OFFICE.NamespaceName, "1.2");
                    w.WriteStartElement("office", "meta", OFFICE.NamespaceName);
                    w.WriteElementString("meta", "generator", META.NamespaceName, "Mini Word 1.1");
                    if (!string.IsNullOrEmpty(doc.Title))
                        w.WriteElementString("dc", "title", DC.NamespaceName, doc.Title);
                    if (!string.IsNullOrEmpty(doc.Author))
                        w.WriteElementString("dc", "creator", DC.NamespaceName, doc.Author);
                    w.WriteElementString("meta", "creation-date", META.NamespaceName,
                        DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildSettingsXml()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("office", "document-settings", OFFICE.NamespaceName);
                    WriteNamespaces(w);
                    w.WriteAttributeString("office", "version", OFFICE.NamespaceName, "1.2");
                    w.WriteStartElement("office", "settings", OFFICE.NamespaceName);
                    w.WriteStartElement("config", "config-item-set", CONFIG.NamespaceName);
                    w.WriteAttributeString("config", "name", CONFIG.NamespaceName, "ooo:view-settings");
                    w.WriteStartElement("config", "config-item", CONFIG.NamespaceName);
                    w.WriteAttributeString("config", "name", CONFIG.NamespaceName, "ViewAreaTop");
                    w.WriteAttributeString("config", "type", CONFIG.NamespaceName, "long");
                    w.WriteString("0");
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildManifest(List<ImageInfo> images)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("manifest", "manifest", MANIFEST.NamespaceName);
                    w.WriteAttributeString("manifest", "version", MANIFEST.NamespaceName, "1.2");
                    WriteManifestEntry(w, "/", "application/vnd.oasis.opendocument.text");
                    WriteManifestEntry(w, "content.xml", "text/xml");
                    WriteManifestEntry(w, "styles.xml", "text/xml");
                    WriteManifestEntry(w, "meta.xml", "text/xml");
                    WriteManifestEntry(w, "settings.xml", "text/xml");
                    for (int i = 0; i < images.Count; i++)
                        WriteManifestEntry(w, "Pictures/" + images[i].FileName, MediaTypeFor(images[i].Extension));
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static void WriteManifestEntry(XmlWriter w, string path, string mediaType)
        {
            w.WriteStartElement("manifest", "file-entry", MANIFEST.NamespaceName);
            w.WriteAttributeString("manifest", "full-path", MANIFEST.NamespaceName, path);
            w.WriteAttributeString("manifest", "media-type", MANIFEST.NamespaceName, mediaType);
            w.WriteEndElement();
        }

        private static string Inches(int twips)
        {
            return (twips / 1440.0).ToString("0.###", CultureInfo.InvariantCulture) + "in";
        }

        private static string Points(int halfPoints)
        {
            return (halfPoints / 2.0).ToString("0.#", CultureInfo.InvariantCulture) + "pt";
        }

        private static string Hex(Color c)
        {
            return "#" + c.R.ToString("X2", CultureInfo.InvariantCulture)
                + c.G.ToString("X2", CultureInfo.InvariantCulture)
                + c.B.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static XmlWriter CreateXmlWriter(Stream stream)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = false;
            settings.CloseOutput = false;
            return XmlWriter.Create(stream, settings);
        }

        // ------------------------------------------------------------------
        // Reading
        // ------------------------------------------------------------------

        public static DocumentModel Read(string path)
        {
            Dictionary<string, byte[]> zip = ZipArchiveLite.Read(path);
            byte[] content;
            if (!zip.TryGetValue("content.xml", out content))
                throw new InvalidDataException("El archivo ODT no contiene content.xml.");

            XDocument contentDoc = LoadXml(content);
            XDocument stylesDoc = null;
            byte[] temp;
            if (zip.TryGetValue("styles.xml", out temp)) stylesDoc = LoadXml(temp);

            OdtReader reader = new OdtReader(zip);
            if (stylesDoc != null) reader.ParseStyleDocument(stylesDoc);
            reader.ParseStyleDocument(contentDoc);
            DocumentModel model = reader.ReadContent(contentDoc);
            bool basic = false;
            bool scripts = false;
            foreach (string key in zip.Keys)
            {
                if (key.StartsWith("Basic/", StringComparison.OrdinalIgnoreCase))
                    basic = true;
                if (key.StartsWith("Scripts/", StringComparison.OrdinalIgnoreCase))
                    scripts = true;
            }
            if (basic || scripts)
                model.SecurityWarnings.Add("El ODT contiene macros o scripts. No se ejecutan y no se conservan al guardar.");
            return model;
        }

        private static XDocument LoadXml(byte[] data)
        {
            if (data == null || data.Length > 64 * 1024 * 1024)
                throw new InvalidDataException("XML del ODT demasiado grande.");
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Prohibit;
            settings.XmlResolver = null;
            settings.MaxCharactersInDocument = 64 * 1024 * 1024;
            settings.MaxCharactersFromEntities = 0;
            using (MemoryStream ms = new MemoryStream(data, false))
            using (XmlReader reader = XmlReader.Create(ms, settings))
            {
                XDocument doc = XDocument.Load(reader);
                return doc;
            }
        }

        private sealed class OdtStyle
        {
            public string Parent;
            public readonly Dictionary<string, string> TextProps = new Dictionary<string, string>(StringComparer.Ordinal);
            public readonly Dictionary<string, string> ParaProps = new Dictionary<string, string>(StringComparer.Ordinal);
        }

        private sealed class ListLevel
        {
            public bool Bullet = true;
            public string NumberFormat = "1";
            public int SpaceBefore;
            public int MinLabelWidth = 360;
        }

        private sealed class ListStyle
        {
            public readonly Dictionary<int, ListLevel> Levels = new Dictionary<int, ListLevel>();
        }

        private sealed class OdtReader
        {
            private readonly Dictionary<string, byte[]> _zip;
            private readonly Dictionary<string, OdtStyle> _styles = new Dictionary<string, OdtStyle>(StringComparer.Ordinal);
            private readonly Dictionary<string, string> _fontFaces = new Dictionary<string, string>(StringComparer.Ordinal);
            private readonly Dictionary<string, ListStyle> _listStyles = new Dictionary<string, ListStyle>(StringComparer.Ordinal);
            private readonly RunFormat _defaultRun = new RunFormat();
            private readonly Paragraph _defaultPara = new Paragraph();
            private readonly Dictionary<string, OdtStyle> _resolved = new Dictionary<string, OdtStyle>(StringComparer.Ordinal);
            private readonly Dictionary<string, int[]> _counters = new Dictionary<string, int[]>();
            private int _pageWidthTwips;
            private int _pageHeightTwips;
            private int _marginLeftTwips;
            private int _marginRightTwips;
            private int _marginTopTwips;
            private int _marginBottomTwips;

            public OdtReader(Dictionary<string, byte[]> zip)
            {
                _zip = zip;
                _defaultRun = new RunFormat("Calibri", 22);
                _defaultPara = new Paragraph();
            }

            public void ParseStyleDocument(XDocument doc)
            {
                XElement root = doc.Root;
                if (root == null) return;

                foreach (XElement ff in root.Descendants(STYLE + "font-face"))
                {
                    string name = (string)ff.Attribute(STYLE + "name");
                    string family = (string)ff.Attribute(SVG + "font-family");
                    if (name == null) continue;
                    if (family != null)
                    {
                        family = family.Trim();
                        if (family.Length >= 2 && (family[0] == '\'' || family[0] == '"'))
                            family = family.Substring(1, family.Length - 2);
                        int comma = family.IndexOf(',');
                        if (comma > 0) family = family.Substring(0, comma).Trim();
                    }
                    _fontFaces[name] = family ?? name;
                }

                foreach (XElement def in root.Descendants(STYLE + "default-style"))
                {
                    string family = (string)def.Attribute(STYLE + "family");
                    OdtStyle s = ParseStyleProperties(def);
                    if (family != null)
                    {
                        if (family == "paragraph")
                        {
                            ApplyParagraphProperties(_defaultPara, s.ParaProps);
                            ApplyTextProperties(_defaultRun, s.TextProps);
                        }
                        else if (family == "text")
                        {
                            ApplyTextProperties(_defaultRun, s.TextProps);
                        }
                    }
                }

                foreach (XElement st in root.Descendants(STYLE + "style"))
                {
                    string family = (string)st.Attribute(STYLE + "family");
                    string name = (string)st.Attribute(STYLE + "name");
                    if (name == null || family == null) continue;
                    OdtStyle s = ParseStyleProperties(st);
                    _styles[family + "|" + name] = s;
                }

                foreach (XElement layout in root.Descendants(STYLE + "page-layout"))
                {
                    XElement props = layout.Element(STYLE + "page-layout-properties");
                    if (props == null) continue;
                    int w = ParseLength((string)props.Attribute(FO + "page-width"));
                    int h = ParseLength((string)props.Attribute(FO + "page-height"));
                    if (w > 0) _pageWidthTwips = w;
                    if (h > 0) _pageHeightTwips = h;
                    int v = ParseLength((string)props.Attribute(FO + "margin-left"));
                    if (v > 0) _marginLeftTwips = v;
                    v = ParseLength((string)props.Attribute(FO + "margin-right"));
                    if (v > 0) _marginRightTwips = v;
                    v = ParseLength((string)props.Attribute(FO + "margin-top"));
                    if (v > 0) _marginTopTwips = v;
                    v = ParseLength((string)props.Attribute(FO + "margin-bottom"));
                    if (v > 0) _marginBottomTwips = v;
                    break;
                }

                foreach (XElement ls in root.Descendants(TEXT + "list-style"))
                {
                    string name = (string)ls.Attribute(STYLE + "name");
                    if (name == null) continue;
                    ListStyle style = new ListStyle();
                    foreach (XElement lev in ls.Elements())
                    {
                        if (lev.Name != TEXT + "list-level-style-bullet" && lev.Name != TEXT + "list-level-style-number")
                            continue;
                        int level = 1;
                        int.TryParse((string)lev.Attribute(TEXT + "level"), out level);
                        ListLevel info = new ListLevel();
                        info.Bullet = lev.Name == TEXT + "list-level-style-bullet";
                        if (!info.Bullet)
                        {
                            string fmt = (string)lev.Attribute(STYLE + "num-format");
                            if (!string.IsNullOrEmpty(fmt)) info.NumberFormat = fmt;
                        }
                        XElement props = lev.Element(STYLE + "list-level-properties");
                        if (props != null)
                        {
                            info.SpaceBefore = ParseLength((string)props.Attribute(TEXT + "space-before"));
                            int minWidth = ParseLength((string)props.Attribute(TEXT + "min-label-width"));
                            if (minWidth > 0) info.MinLabelWidth = minWidth;
                        }
                        style.Levels[level] = info;
                    }
                    _listStyles[name] = style;
                }
            }

            private OdtStyle ParseStyleProperties(XElement element)
            {
                OdtStyle s = new OdtStyle();
                s.Parent = (string)element.Attribute(STYLE + "parent-style-name");

                foreach (XElement child in element.Elements())
                {
                    if (child.Name == STYLE + "text-properties")
                    {
                        foreach (XAttribute a in child.Attributes())
                        {
                            string key = a.Name.LocalName;
                            if (a.Name.Namespace == STYLE || a.Name.Namespace == FO)
                                s.TextProps[key] = a.Value;
                        }
                    }
                    else if (child.Name == STYLE + "paragraph-properties")
                    {
                        foreach (XAttribute a in child.Attributes())
                        {
                            string key = a.Name.LocalName;
                            if (a.Name.Namespace == STYLE || a.Name.Namespace == FO)
                                s.ParaProps[key] = a.Value;
                        }
                    }
                }
                return s;
            }

            private OdtStyle ResolveStyle(string name, string family)
            {
                if (string.IsNullOrEmpty(name)) return null;
                string key = family + "|" + name;
                OdtStyle cached;
                if (_resolved.TryGetValue(key, out cached)) return cached;

                OdtStyle own;
                if (!_styles.TryGetValue(key, out own)) own = null;
                if (own == null) return null;

                _resolved[key] = null;
                OdtStyle merged = new OdtStyle();
                if (!string.IsNullOrEmpty(own.Parent))
                {
                    OdtStyle parent = ResolveStyle(own.Parent, family);
                    if (parent != null)
                    {
                        foreach (KeyValuePair<string, string> kv in parent.TextProps) merged.TextProps[kv.Key] = kv.Value;
                        foreach (KeyValuePair<string, string> kv in parent.ParaProps) merged.ParaProps[kv.Key] = kv.Value;
                    }
                }
                foreach (KeyValuePair<string, string> kv in own.TextProps) merged.TextProps[kv.Key] = kv.Value;
                foreach (KeyValuePair<string, string> kv in own.ParaProps) merged.ParaProps[kv.Key] = kv.Value;
                _resolved[key] = merged;
                return merged;
            }

            public DocumentModel ReadContent(XDocument content)
            {
                DocumentModel model = new DocumentModel();
                XElement root = content.Root;
                XElement body = root != null ? root.Element(OFFICE + "body") : null;
                XElement text = body != null ? body.Element(OFFICE + "text") : null;

                if (_pageWidthTwips > 0 && _pageHeightTwips > 0)
                    model.SetPageSize(_pageWidthTwips, _pageHeightTwips);
                if (_marginLeftTwips > 0) model.MarginLeftTwips = _marginLeftTwips;
                if (_marginRightTwips > 0) model.MarginRightTwips = _marginRightTwips;
                if (_marginTopTwips > 0) model.MarginTopTwips = _marginTopTwips;
                if (_marginBottomTwips > 0) model.MarginBottomTwips = _marginBottomTwips;

                if (root != null)
                {
                    foreach (XElement layout in root.Descendants(STYLE + "page-layout"))
                    {
                        XElement props = layout.Element(STYLE + "page-layout-properties");
                        if (props == null) continue;
                        int w = ParseLength((string)props.Attribute(FO + "page-width"));
                        int h = ParseLength((string)props.Attribute(FO + "page-height"));
                        if (w > 0 && h > 0) model.SetPageSize(w, h);
                        int v = ParseLength((string)props.Attribute(FO + "margin-left"));
                        if (v > 0) model.MarginLeftTwips = v;
                        v = ParseLength((string)props.Attribute(FO + "margin-right"));
                        if (v > 0) model.MarginRightTwips = v;
                        v = ParseLength((string)props.Attribute(FO + "margin-top"));
                        if (v > 0) model.MarginTopTwips = v;
                        v = ParseLength((string)props.Attribute(FO + "margin-bottom"));
                        if (v > 0) model.MarginBottomTwips = v;
                        break;
                    }
                }

                if (text != null) ReadTextChildren(text, model, 0, 0);

                if (model.Paragraphs.Count == 0) model.Paragraphs.Add(new Paragraph());
                return model;
            }

            private void ReadTextChildren(XElement parent, DocumentModel model, int listLevel, int depth)
            {
                if (depth > 128) throw new InvalidDataException("El ODT contiene un anidamiento excesivo.");
                foreach (XElement child in parent.Elements())
                {
                    if (child.Name == TEXT + "p")
                    {
                        Paragraph p = new Paragraph();
                        CopyDefaultParagraph(p);
                        ReadParagraph(child, p, _defaultRun.Clone());
                        model.Paragraphs.Add(p);
                    }
                    else if (child.Name == TEXT + "h")
                    {
                        Paragraph p = new Paragraph();
                        CopyDefaultParagraph(p);
                        RunFormat fmt = _defaultRun.Clone();
                        int level = 1;
                        int.TryParse((string)child.Attribute(TEXT + "outline-level"), out level);
                        fmt.Bold = true;
                        fmt.SizeHalfPoints = level <= 1 ? 32 : (level == 2 ? 28 : 24);
                        ReadParagraph(child, p, fmt);
                        model.Paragraphs.Add(p);
                    }
                    else if (child.Name == TEXT + "list")
                    {
                        ReadList(child, model, listLevel, depth + 1);
                    }
                    else if (child.Name == TABLE + "table")
                    {
                        ReadTable(child, model);
                    }
                    else if (child.Name == TEXT + "section")
                    {
                        ReadTextChildren(child, model, listLevel, depth + 1);
                    }
                }
            }

            private void ReadList(XElement list, DocumentModel model, int parentLevel, int depth)
            {
                if (depth > 128) throw new InvalidDataException("El ODT contiene un anidamiento excesivo.");
                string styleName = (string)list.Attribute(TEXT + "style-name");
                ListStyle style = null;
                if (styleName != null) _listStyles.TryGetValue(styleName, out style);
                string counterKey = styleName + "|" + list.GetHashCode().ToString(CultureInfo.InvariantCulture);
                int[] counters;
                if (!_counters.TryGetValue(counterKey, out counters))
                {
                    counters = new int[12];
                    _counters[counterKey] = counters;
                }
                int level = parentLevel + 1;

                foreach (XElement item in list.Elements(TEXT + "list-item"))
                {
                    bool first = true;
                    foreach (XElement child in item.Elements())
                    {
                        if (child.Name == TEXT + "p" || child.Name == TEXT + "h")
                        {
                            Paragraph p = new Paragraph();
                            CopyDefaultParagraph(p);
                            RunFormat fmt = _defaultRun.Clone();
                            ReadParagraph(child, p, fmt);

                            ListLevel info = null;
                            if (style != null) style.Levels.TryGetValue(level, out info);
                            if (info == null) info = new ListLevel();

                            if (first)
                            {
                                if (info.Bullet)
                                {
                                    p.Bullet = true;
                                    p.BulletLevel = Math.Max(0, level - 1);
                                }
                                else
                                {
                                    if (level < counters.Length) counters[level]++;
                                    for (int k = level + 1; k < counters.Length; k++) counters[k] = 0;
                                    string prefix = FormatListNumber(counters[level], info.NumberFormat) + " ";
                                    p.Inlines.Insert(0, new TextInline(prefix, fmt.Clone()));
                                }
                                int left = info.SpaceBefore + info.MinLabelWidth;
                                if (left <= 0) left = 720 * level;
                                if (p.LeftIndentTwips == 0) p.LeftIndentTwips = left;
                                if (p.FirstLineIndentTwips == 0) p.FirstLineIndentTwips = -info.MinLabelWidth;
                            }
                            else
                            {
                                if (p.LeftIndentTwips == 0) p.LeftIndentTwips = 720 * level;
                            }
                            model.Paragraphs.Add(p);
                            first = false;
                        }
                        else if (child.Name == TEXT + "list")
                        {
                            ReadList(child, model, level, depth + 1);
                        }
                    }
                }
            }

            private void ReadParagraph(XElement element, Paragraph p, RunFormat baseRun)
            {
                string styleName = (string)element.Attribute(TEXT + "style-name");
                if (!string.IsNullOrEmpty(styleName))
                {
                    OdtStyle style = ResolveStyle(styleName, "paragraph");
                    if (style != null)
                    {
                        ApplyParagraphProperties(p, style.ParaProps);
                        ApplyTextProperties(baseRun, style.TextProps);
                    }
                }

                ReadInline(element, p, baseRun, 0);
            }

            private void ReadInline(XElement parent, Paragraph p, RunFormat fmt, int depth)
            {
                if (depth > 128) throw new InvalidDataException("El ODT contiene un anidamiento excesivo.");
                foreach (XNode node in parent.Nodes())
                {
                    XText text = node as XText;
                    if (text != null)
                    {
                        string value = text.Value;
                        if (value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0)
                            value = Regex.Replace(value, "[ \t\r\n]+", " ");
                        if (value.Length > 0) AppendText(p, value, fmt);
                        continue;
                    }

                    XElement child = node as XElement;
                    if (child == null) continue;

                    if (child.Name == TEXT + "span")
                    {
                        RunFormat childFmt = fmt.Clone();
                        string styleName = (string)child.Attribute(TEXT + "style-name");
                        if (!string.IsNullOrEmpty(styleName))
                        {
                            OdtStyle style = ResolveStyle(styleName, "text");
                            if (style != null) ApplyTextProperties(childFmt, style.TextProps);
                        }
                        ReadInline(child, p, childFmt, depth + 1);
                    }
                    else if (child.Name == TEXT + "a")
                    {
                        RunFormat link = fmt.Clone();
                        link.ForeColor = Color.FromArgb(0, 0, 238);
                        link.Underline = true;
                        link.LinkUrl = (string)child.Attribute(XLINK + "href");
                        ReadInline(child, p, link, depth + 1);
                    }
                    else if (child.Name == TEXT + "s")
                    {
                        int count = 1;
                        int.TryParse((string)child.Attribute(TEXT + "c"), out count);
                        if (count < 1) count = 1;
                        AppendText(p, new string(' ', count), fmt);
                    }
                    else if (child.Name == TEXT + "tab")
                    {
                        AppendText(p, "\t", fmt);
                    }
                    else if (child.Name == TEXT + "line-break")
                    {
                        AppendText(p, "\n", fmt);
                    }
                    else if (child.Name == DRAW + "frame" || child.Name == DRAW + "custom-shape")
                    {
                        ImageInline img = ReadFrame(child);
                        if (img != null)
                        {
                            img.Format = fmt.Clone();
                            p.Inlines.Add(img);
                        }
                    }
                    else if (child.Name == TEXT + "bookmark" || child.Name == TEXT + "bookmark-start"
                        || child.Name == TEXT + "bookmark-end" || child.Name == TEXT + "soft-page-break")
                    {
                        // ignored
                    }
                    else if (child.Name == TEXT + "note")
                    {
                        // footnote/endnote bodies are not part of the main text flow
                    }
                    else if (child.Name == OFFICE + "annotation")
                    {
                        // ignored
                    }
                    else
                    {
                        ReadInline(child, p, fmt, depth + 1);
                    }
                }
            }

            private ImageInline ReadFrame(XElement frame)
            {
                XElement image = frame.Descendants(DRAW + "image").FirstOrDefault();
                if (image == null) return null;
                string href = (string)image.Attribute(XLINK + "href");
                if (string.IsNullOrEmpty(href)) return null;
                if (href.StartsWith("./", StringComparison.Ordinal)) href = href.Substring(2);
                byte[] data;
                if (!_zip.TryGetValue(href, out data)) return null;
                int w = ParseLength((string)frame.Attribute(SVG + "width"));
                int h = ParseLength((string)frame.Attribute(SVG + "height"));
                if (w <= 0) w = 1440;
                if (h <= 0) h = 1440;
                string ext = "png";
                int dot = href.LastIndexOf('.');
                if (dot >= 0) ext = href.Substring(dot + 1).ToLowerInvariant();
                if (ext == "jpg") ext = "jpeg";
                return new ImageInline(data, ext, w, h, new RunFormat());
            }

            private void ReadTable(XElement table, DocumentModel model)
            {
                foreach (XElement row in table.Elements(TABLE + "table-row"))
                {
                    List<string> cells = new List<string>();
                    foreach (XElement cell in row.Elements(TABLE + "table-cell"))
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (XElement para in cell.Descendants(TEXT + "p"))
                        {
                            if (sb.Length > 0) sb.Append(' ');
                            foreach (XNode node in para.DescendantNodes())
                            {
                                XText t = node as XText;
                                if (t != null) sb.Append(t.Value);
                            }
                        }
                        cells.Add(sb.ToString());
                    }
                    Paragraph p = new Paragraph();
                    CopyDefaultParagraph(p);
                    p.Inlines.Add(new TextInline(string.Join("\t", cells.ToArray()), _defaultRun.Clone()));
                    model.Paragraphs.Add(p);
                }
            }

            private void CopyDefaultParagraph(Paragraph p)
            {
                p.Alignment = _defaultPara.Alignment;
                p.LeftIndentTwips = _defaultPara.LeftIndentTwips;
                p.RightIndentTwips = _defaultPara.RightIndentTwips;
                p.FirstLineIndentTwips = _defaultPara.FirstLineIndentTwips;
                p.SpaceBeforeTwips = _defaultPara.SpaceBeforeTwips;
                p.SpaceAfterTwips = _defaultPara.SpaceAfterTwips;
                p.LineSpacing = _defaultPara.LineSpacing;
            }

            private void ApplyParagraphProperties(Paragraph p, Dictionary<string, string> props)
            {
                string v;
                if (props.TryGetValue("text-align", out v))
                {
                    if (v == "center") p.Alignment = ParagraphAlignment.Center;
                    else if (v == "end" || v == "right") p.Alignment = ParagraphAlignment.Right;
                    else if (v == "justify") p.Alignment = ParagraphAlignment.Justify;
                    else if (v == "start" || v == "left") p.Alignment = ParagraphAlignment.Left;
                }
                if (props.TryGetValue("margin-left", out v)) p.LeftIndentTwips = ParseLength(v);
                if (props.TryGetValue("margin-right", out v)) p.RightIndentTwips = ParseLength(v);
                if (props.TryGetValue("text-indent", out v)) p.FirstLineIndentTwips = ParseLength(v);
                if (props.TryGetValue("margin-top", out v)) p.SpaceBeforeTwips = ParseLength(v);
                if (props.TryGetValue("margin-bottom", out v)) p.SpaceAfterTwips = ParseLength(v);
                if (props.TryGetValue("line-height", out v))
                {
                    double mult = ParseLineHeight(v);
                    if (mult >= 1.8) p.LineSpacing = 2;
                    else if (mult >= 1.3) p.LineSpacing = 1;
                    else p.LineSpacing = 0;
                }
            }

            private void ApplyTextProperties(RunFormat f, Dictionary<string, string> props)
            {
                string v;
                if (props.TryGetValue("font-name", out v))
                {
                    string family;
                    if (_fontFaces.TryGetValue(v, out family)) f.FontFamily = family;
                    else f.FontFamily = v;
                }
                else if (props.TryGetValue("font-family", out v))
                {
                    f.FontFamily = CleanFamily(v);
                }
                if (props.TryGetValue("font-size", out v))
                {
                    int hp = ParseFontSize(v);
                    if (hp > 0) f.SizeHalfPoints = hp;
                }
                if (props.TryGetValue("font-weight", out v))
                    f.Bold = v == "bold" || IsNumericAtLeast(v, 600);
                if (props.TryGetValue("font-style", out v))
                    f.Italic = v == "italic" || v == "oblique";
                if (props.TryGetValue("color", out v))
                {
                    Color c;
                    if (TryColor(v, out c)) f.ForeColor = c;
                    else if (v == "transparent") f.ForeColor = null;
                }
                if (props.TryGetValue("background-color", out v))
                {
                    Color c;
                    if (TryColor(v, out c)) f.BackColor = c;
                    else if (v == "transparent" || v == "none") f.BackColor = null;
                }
                if (props.TryGetValue("text-underline-style", out v))
                    f.Underline = v != "none" && v != "0";
                if (props.TryGetValue("text-line-through-style", out v))
                    f.Strikeout = v != "none" && v != "0";
                if (props.TryGetValue("text-position", out v))
                {
                    if (v.StartsWith("super", StringComparison.OrdinalIgnoreCase)) f.Script = 1;
                    else if (v.StartsWith("sub", StringComparison.OrdinalIgnoreCase)) f.Script = -1;
                }
            }

            private static string CleanFamily(string family)
            {
                if (string.IsNullOrEmpty(family)) return family;
                family = family.Trim();
                if (family.Length >= 2 && (family[0] == '\'' || family[0] == '"'))
                    family = family.Substring(1, family.Length - 2);
                int comma = family.IndexOf(',');
                if (comma > 0) family = family.Substring(0, comma).Trim();
                return family;
            }

            private static bool IsNumericAtLeast(string v, int min)
            {
                int n;
                return int.TryParse(v, out n) && n >= min;
            }

            private static bool TryColor(string v, out Color color)
            {
                color = Color.Black;
                if (string.IsNullOrEmpty(v)) return false;
                v = v.Trim();
                if (!v.StartsWith("#", StringComparison.Ordinal)) return false;
                v = v.Substring(1);
                if (v.Length == 3)
                    v = new string(new char[] { v[0], v[0], v[1], v[1], v[2], v[2] });
                if (v.Length != 6) return false;
                int r, g, b;
                if (!int.TryParse(v.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r)) return false;
                if (!int.TryParse(v.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g)) return false;
                if (!int.TryParse(v.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b)) return false;
                color = Color.FromArgb(r, g, b);
                return true;
            }

            private static void AppendText(Paragraph p, string text, RunFormat fmt)
            {
                if (string.IsNullOrEmpty(text)) return;
                if (p.Inlines.Count > 0)
                {
                    TextInline last = p.Inlines[p.Inlines.Count - 1] as TextInline;
                    if (last != null && last.Format.SameAs(fmt))
                    {
                        last.Text += text;
                        return;
                    }
                }
                p.Inlines.Add(new TextInline(text, fmt.Clone()));
            }

            private static string FormatListNumber(int n, string format)
            {
                switch (format)
                {
                    case "A": return ToLetters(n, true);
                    case "I": return ToRoman(n, true);
                }
                switch ((format ?? "1").ToLowerInvariant())
                {
                    case "a": return ToLetters(n, false);
                    case "i": return ToRoman(n, false);
                    default: return n.ToString(CultureInfo.InvariantCulture);
                }
            }

            private static string ToLetters(int n, bool upper)
            {
                if (n <= 0) n = 1;
                StringBuilder sb = new StringBuilder();
                while (n > 0)
                {
                    n--;
                    sb.Insert(0, (char)('a' + n % 26));
                    n /= 26;
                }
                string s = sb.ToString();
                return upper ? s.ToUpperInvariant() : s;
            }

            private static string ToRoman(int n, bool upper)
            {
                if (n <= 0) n = 1;
                int[] values = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
                string[] symbols = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < values.Length; i++)
                {
                    while (n >= values[i])
                    {
                        sb.Append(symbols[i]);
                        n -= values[i];
                    }
                }
                string s = sb.ToString();
                return upper ? s : s.ToLowerInvariant();
            }
        }

        private static int ParseLength(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            value = value.Trim();
            try
            {
                if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
                    return (int)Math.Round(double.Parse(value.Substring(0, value.Length - 2), CultureInfo.InvariantCulture) * 1440);
                if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
                    return (int)Math.Round(double.Parse(value.Substring(0, value.Length - 2), CultureInfo.InvariantCulture) * 566.929);
                if (value.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
                    return (int)Math.Round(double.Parse(value.Substring(0, value.Length - 2), CultureInfo.InvariantCulture) * 56.6929);
                if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                    return (int)Math.Round(double.Parse(value.Substring(0, value.Length - 2), CultureInfo.InvariantCulture) * 20);
                if (value.EndsWith("pc", StringComparison.OrdinalIgnoreCase))
                    return (int)Math.Round(double.Parse(value.Substring(0, value.Length - 2), CultureInfo.InvariantCulture) * 240);
                if (value.EndsWith("px", StringComparison.OrdinalIgnoreCase))
                    return (int)Math.Round(double.Parse(value.Substring(0, value.Length - 2), CultureInfo.InvariantCulture) * 15);
                double plain;
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out plain))
                    return (int)Math.Round(plain * 20);
            }
            catch
            {
            }
            return 0;
        }

        private static int ParseFontSize(string value)
        {
            int twips = ParseLength(value);
            if (twips <= 0) return 0;
            return Math.Max(1, (int)Math.Round(twips / 10.0));
        }

        private static double ParseLineHeight(string value)
        {
            if (string.IsNullOrEmpty(value)) return 1;
            value = value.Trim();
            try
            {
                if (value.EndsWith("%", StringComparison.Ordinal))
                    return double.Parse(value.Substring(0, value.Length - 1), CultureInfo.InvariantCulture) / 100.0;
                if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase) || value.EndsWith("cm", StringComparison.OrdinalIgnoreCase)
                    || value.EndsWith("mm", StringComparison.OrdinalIgnoreCase) || value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                {
                    int twips = ParseLength(value);
                    return twips / 240.0;
                }
                double d;
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
            }
            catch
            {
            }
            return 1;
        }
    }
}
