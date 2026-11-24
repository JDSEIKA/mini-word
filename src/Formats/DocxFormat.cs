using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace WordPadClone.Formats
{
    /// <summary>
    /// Office Open XML (WordprocessingML) reader and writer.  It implements the
    /// subset WordPad itself handles, plus pictures, bullets and page setup.
    /// </summary>
    public static class DocxFormat
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace VML = "urn:schemas-microsoft-com:vml";
        private static readonly XNamespace WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        private static readonly XNamespace OLE = "urn:schemas-microsoft-com:office:office";
        private static readonly XNamespace REL = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace CT = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static readonly XNamespace DC = "http://purl.org/dc/elements/1.1/";
        private static readonly XNamespace DCTERMS = "http://purl.org/dc/terms/";
        private static readonly XNamespace XSI = "http://www.w3.org/2001/XMLSchema-instance";
        private static readonly XNamespace EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

        // ------------------------------------------------------------------
        // Writing
        // ------------------------------------------------------------------

        public static void Write(string path, DocumentModel doc)
        {
            List<ZipEntryData> entries = new List<ZipEntryData>();
            List<ImagePart> images = new List<ImagePart>();
            bool hasBullets = false;
            bool hasNumbered = false;
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                if (doc.Paragraphs[i].Bullet) hasBullets = true;
                if (doc.Paragraphs[i].Numbered) hasNumbered = true;
            }
            bool hasLists = hasBullets || hasNumbered;

            Dictionary<string, int> imageIds = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Paragraph p = doc.Paragraphs[i];
                for (int j = 0; j < p.Inlines.Count; j++)
                    AddImagePart(p.Inlines[j] as ImageInline, images, imageIds);
                for (int j = 0; j < p.FloatingObjects.Count; j++)
                    if (p.FloatingObjects[j].Image != null) AddImagePart(p.FloatingObjects[j].Image, images, imageIds);
            }

            List<EmbeddedPart> embedded = new List<EmbeddedPart>();
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Paragraph p = doc.Paragraphs[i];
                for (int j = 0; j < p.FloatingObjects.Count; j++)
                {
                    FloatingObject obj = p.FloatingObjects[j];
                    if (obj.Kind != FloatingKind.Embedded || obj.EmbeddedData == null || obj.EmbeddedData.Length == 0) continue;
                    bool duplicate = false;
                    for (int k = 0; k < embedded.Count; k++)
                        if (embedded[k].Data.Length == obj.EmbeddedData.Length && SameBytes(embedded[k].Data, obj.EmbeddedData)) duplicate = true;
                    if (duplicate) continue;
                    EmbeddedPart part = new EmbeddedPart();
                    part.RelId = "rIdOle" + (embedded.Count + 1);
                    part.FileName = "oleObject" + (embedded.Count + 1) + ".bin";
                    part.Data = obj.EmbeddedData;
                    embedded.Add(part);
                }
            }

            doc.EnsureSections();
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                for (int j = 0; j < doc.Sections[i].HeaderImages.Count; j++)
                    AddImagePart(doc.Sections[i].HeaderImages[j], images, imageIds);
                for (int j = 0; j < doc.Sections[i].FooterImages.Count; j++)
                    AddImagePart(doc.Sections[i].FooterImages[j], images, imageIds);
                for (int j = 0; j < doc.Sections[i].FirstPageHeaderImages.Count; j++)
                    AddImagePart(doc.Sections[i].FirstPageHeaderImages[j], images, imageIds);
                for (int j = 0; j < doc.Sections[i].FirstPageFooterImages.Count; j++)
                    AddImagePart(doc.Sections[i].FirstPageFooterImages[j], images, imageIds);
                for (int j = 0; j < doc.Sections[i].EvenPageHeaderImages.Count; j++)
                    AddImagePart(doc.Sections[i].EvenPageHeaderImages[j], images, imageIds);
                for (int j = 0; j < doc.Sections[i].EvenPageFooterImages.Count; j++)
                    AddImagePart(doc.Sections[i].EvenPageFooterImages[j], images, imageIds);
            }
            List<HeaderFooterPart> headers = BuildHeaderFooterParts(doc, true);
            List<HeaderFooterPart> footers = BuildHeaderFooterParts(doc, false);
            bool hasFootnotes = doc.Footnotes.Count > 0;
            bool hasEndnotes = doc.Endnotes.Count > 0;
            bool hasComments = doc.Comments.Count > 0;
            List<HyperlinkPart> hyperlinks = new List<HyperlinkPart>();
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Paragraph p = doc.Paragraphs[i];
                for (int j = 0; j < p.Inlines.Count; j++)
                {
                    TextInline text = p.Inlines[j] as TextInline;
                    if (text == null || string.IsNullOrEmpty(text.Format.LinkUrl)) continue;
                    bool exists = false;
                    for (int k = 0; k < hyperlinks.Count; k++)
                        if (string.Equals(hyperlinks[k].Url, text.Format.LinkUrl, StringComparison.Ordinal)) { exists = true; break; }
                    if (exists) continue;
                    HyperlinkPart link = new HyperlinkPart();
                    link.RelId = "rIdLink" + (hyperlinks.Count + 1);
                    link.Url = text.Format.LinkUrl;
                    hyperlinks.Add(link);
                }
                p.EnsureTableCells();
                if (p.TableCells != null)
                {
                    for (int r = 0; r < p.TableCells.Length; r++)
                    {
                        if (p.TableCells[r] == null) continue;
                        for (int c = 0; c < p.TableCells[r].Length; c++)
                        {
                            TableCell cell = p.TableCells[r][c];
                            if (cell == null) continue;
                            if (cell.Runs.Count > 0)
                            {
                                for (int q = 0; q < cell.Runs.Count; q++)
                                {
                                    string runUrl = cell.Runs[q].LinkUrl;
                                    if (string.IsNullOrEmpty(runUrl)) continue;
                                    bool runExists = false;
                                    for (int k = 0; k < hyperlinks.Count; k++)
                                        if (string.Equals(hyperlinks[k].Url, runUrl, StringComparison.Ordinal)) { runExists = true; break; }
                                    if (!runExists)
                                    {
                                        HyperlinkPart runLink = new HyperlinkPart();
                                        runLink.RelId = "rIdLink" + (hyperlinks.Count + 1);
                                        runLink.Url = runUrl;
                                        hyperlinks.Add(runLink);
                                    }
                                }
                            }
                            if (string.IsNullOrEmpty(cell.LinkUrl)) continue;
                            bool exists = false;
                            for (int k = 0; k < hyperlinks.Count; k++)
                                if (string.Equals(hyperlinks[k].Url, cell.LinkUrl, StringComparison.Ordinal)) { exists = true; break; }
                            if (exists) continue;
                            HyperlinkPart link = new HyperlinkPart();
                            link.RelId = "rIdLink" + (hyperlinks.Count + 1);
                            link.Url = cell.LinkUrl;
                            hyperlinks.Add(link);
                        }
                    }
                }
            }
            entries.Add(new ZipEntryData("[Content_Types].xml", BuildContentTypes(doc, images, hasLists, headers, footers, hasFootnotes, hasEndnotes, hasComments, embedded), true));
            entries.Add(new ZipEntryData("_rels/.rels", BuildRootRels(), true));
            entries.Add(new ZipEntryData("word/document.xml", BuildDocumentXml(doc, images, headers, footers, embedded, hyperlinks), true));
            entries.Add(new ZipEntryData("word/styles.xml", BuildStylesXml(), true));
            entries.Add(new ZipEntryData("word/settings.xml", BuildSettingsXml(doc), true));
            if (hasLists)
                entries.Add(new ZipEntryData("word/numbering.xml", BuildNumberingXml(), true));
            for (int i = 0; i < headers.Count; i++)
            {
                entries.Add(new ZipEntryData("word/" + headers[i].FileName,
                    BuildHeaderFooterXml(headers[i], images), true));
                entries.Add(new ZipEntryData("word/_rels/" + headers[i].FileName + ".rels",
                    BuildHeaderFooterRels(headers[i], images), true));
            }
            for (int i = 0; i < footers.Count; i++)
            {
                entries.Add(new ZipEntryData("word/" + footers[i].FileName,
                    BuildHeaderFooterXml(footers[i], images), true));
                entries.Add(new ZipEntryData("word/_rels/" + footers[i].FileName + ".rels",
                    BuildHeaderFooterRels(footers[i], images), true));
            }
            if (hasFootnotes) entries.Add(new ZipEntryData("word/footnotes.xml", BuildNotesXml(doc.Footnotes, false), true));
            if (hasEndnotes) entries.Add(new ZipEntryData("word/endnotes.xml", BuildNotesXml(doc.Endnotes, true), true));
            if (hasComments) entries.Add(new ZipEntryData("word/comments.xml", BuildCommentsXml(doc), true));
            entries.Add(new ZipEntryData("word/_rels/document.xml.rels",
                BuildDocumentRels(doc, images, hasLists, headers, footers, hasFootnotes, hasEndnotes, hasComments, embedded, hyperlinks), true));
            entries.Add(new ZipEntryData("docProps/core.xml", BuildCoreProps(doc), true));
            entries.Add(new ZipEntryData("docProps/app.xml", BuildAppProps(), true));
            for (int i = 0; i < images.Count; i++)
                entries.Add(new ZipEntryData("word/media/" + images[i].FileName, images[i].Data, true));
            for (int i = 0; i < embedded.Count; i++)
                entries.Add(new ZipEntryData("word/embeddings/" + embedded[i].FileName, embedded[i].Data, true));
            AppendPreservedParts(doc, entries);

            ZipArchiveLite.Write(path, entries);
        }

        private static void AppendPreservedParts(DocumentModel doc, List<ZipEntryData> entries)
        {
            if (doc == null || doc.PreservedParts == null || entries == null) return;
            foreach (KeyValuePair<string, byte[]> part in doc.PreservedParts)
            {
                if (part.Value == null) continue;
                bool exists = false;
                for (int i = 0; i < entries.Count; i++)
                {
                    if (string.Equals(entries[i].Name, part.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        exists = true;
                        break;
                    }
                }
                if (!exists) entries.Add(new ZipEntryData(part.Key, part.Value, true));
            }
        }

        private static List<HeaderFooterPart> BuildHeaderFooterParts(DocumentModel doc, bool header)
        {
            List<HeaderFooterPart> result = new List<HeaderFooterPart>();
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                DocumentSection section = doc.Sections[i];
                string text = header ? section.HeaderText : section.FooterText;
                if (string.IsNullOrEmpty(text) && i == 0) text = header ? doc.HeaderText : doc.FooterText;
                List<ImageInline> images = header ? section.HeaderImages : section.FooterImages;
                AddHeaderFooterPart(result, section, header, "default", text, images,
                    header ? section.HeaderFormat : section.FooterFormat,
                    header ? section.HeaderAlignment : section.FooterAlignment);
                if (section.TitlePage && section.FirstPageHeaderText != null)
                    AddHeaderFooterPart(result, section, header, "first",
                        header ? section.FirstPageHeaderText : section.FirstPageFooterText,
                        header ? section.FirstPageHeaderImages : section.FirstPageFooterImages,
                        header ? section.FirstPageHeaderFormat : section.FirstPageFooterFormat,
                        header ? section.FirstPageHeaderAlignment : section.FirstPageFooterAlignment);
                if (section.DifferentOddEven)
                    AddHeaderFooterPart(result, section, header, "even",
                        header ? section.EvenPageHeaderText : section.EvenPageFooterText,
                        header ? section.EvenPageHeaderImages : section.EvenPageFooterImages,
                        header ? section.EvenPageHeaderFormat : section.EvenPageFooterFormat,
                        header ? section.EvenPageHeaderAlignment : section.EvenPageFooterAlignment);
            }
            return result;
        }

        private static void AddHeaderFooterPart(List<HeaderFooterPart> result, DocumentSection section,
            bool header, string type, string text, List<ImageInline> images,
            RunFormat format, ParagraphAlignment alignment)
        {
            if (string.IsNullOrEmpty(text) && (images == null || images.Count == 0)) return;
            HeaderFooterPart part = new HeaderFooterPart();
            part.Section = section;
            part.IsHeader = header;
            part.Type = type;
            part.Text = text ?? string.Empty;
            part.Format = format ?? (header ? section.HeaderFormat : section.FooterFormat) ?? new RunFormat("Calibri", 18);
            part.Alignment = alignment;
            part.Images = images ?? new List<ImageInline>();
            part.FileName = (header ? "header" : "footer") + (result.Count + 1) + ".xml";
            part.RelId = "rId" + (header ? "Hdr" : "Ftr") + (result.Count + 1);
            result.Add(part);
        }

        private static HeaderFooterPart FindHeaderFooter(List<HeaderFooterPart> parts, DocumentSection section,
            string type)
        {
            if (parts == null || section == null) return null;
            for (int i = 0; i < parts.Count; i++)
                if (object.ReferenceEquals(parts[i].Section, section) &&
                    string.Equals(parts[i].Type, type, StringComparison.OrdinalIgnoreCase))
                    return parts[i];
            return null;
        }

        private static void AddImagePart(ImageInline img, List<ImagePart> images, Dictionary<string, int> imageIds)
        {
            if (img == null || img.Data == null || img.Data.Length == 0) return;
            string key = img.Extension + ":" + img.Data.Length + ":" + img.WidthTwips + "x" + img.HeightTwips;
            if (imageIds.ContainsKey(key)) return;
            ImagePart part = new ImagePart();
            part.RelId = "rIdImg" + (images.Count + 1);
            part.FileName = "image" + (images.Count + 1) + "." + NormalizeExtension(img.Extension);
            part.Data = img.Data;
            part.Extension = NormalizeExtension(img.Extension);
            images.Add(part);
            imageIds[key] = images.Count - 1;
        }

        private sealed class ImagePart
        {
            public string RelId;
            public string FileName;
            public string Extension;
            public byte[] Data;
        }

        private sealed class EmbeddedPart
        {
            public string RelId;
            public string FileName;
            public byte[] Data;
        }

        private sealed class HyperlinkPart
        {
            public string RelId;
            public string Url;
        }

        private sealed class HeaderFooterPart
        {
            public DocumentSection Section;
            public bool IsHeader;
            public string Type = "default";
            public string Text;
            public RunFormat Format;
            public ParagraphAlignment Alignment;
            public List<ImageInline> Images;
            public string FileName;
            public string RelId;
        }

        private static string NormalizeExtension(string ext)
        {
            if (string.IsNullOrEmpty(ext)) return "png";
            ext = ext.ToLowerInvariant();
            if (ext == "jpg") ext = "jpeg";
            if (ext != "png" && ext != "jpeg" && ext != "gif" && ext != "bmp" && ext != "emf" && ext != "wmf")
                ext = "png";
            return ext;
        }

        private static string ContentTypeFor(string ext)
        {
            switch (ext)
            {
                case "jpeg": return "image/jpeg";
                case "gif": return "image/gif";
                case "bmp": return "image/bmp";
                case "emf": return "image/x-emf";
                case "wmf": return "image/x-wmf";
                default: return "image/png";
            }
        }

        private static byte[] BuildContentTypes(DocumentModel doc, List<ImagePart> images, bool hasBullets,
            List<HeaderFooterPart> headers, List<HeaderFooterPart> footers,
            bool hasFootnotes, bool hasEndnotes, bool hasComments, List<EmbeddedPart> embedded)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("Types", CT.NamespaceName);
                    WriteDefault(w, "rels", "application/vnd.openxmlformats-package.relationships+xml");
                    WriteDefault(w, "xml", "application/xml");
                    if (embedded.Count > 0)
                        WriteDefault(w, "bin", "application/vnd.openxmlformats-officedocument.oleObject");
                    List<string> exts = new List<string>();
                    for (int i = 0; i < images.Count; i++)
                        if (!exts.Contains(images[i].Extension)) exts.Add(images[i].Extension);
                    for (int i = 0; i < exts.Count; i++)
                        WriteDefault(w, exts[i], ContentTypeFor(exts[i]));
                    if (doc != null && doc.PreservedContentTypes != null)
                    {
                        for (int i = 0; i < doc.PreservedContentTypes.Count; i++)
                        {
                            PreservedContentType preserved = doc.PreservedContentTypes[i];
                            if (string.IsNullOrEmpty(preserved.Extension)) continue;
                            string ext = preserved.Extension.TrimStart('.').ToLowerInvariant();
                            if (ext == "rels" || ext == "xml" || ext == "bin" || exts.Contains(ext)) continue;
                            WriteDefault(w, ext, preserved.ContentType);
                            exts.Add(ext);
                        }
                    }
                    WriteOverride(w, "/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
                    WriteOverride(w, "/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
                    WriteOverride(w, "/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");
                    if (hasBullets)
                        WriteOverride(w, "/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
                    for (int i = 0; i < headers.Count; i++)
                        WriteOverride(w, "/word/" + headers[i].FileName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
                    for (int i = 0; i < footers.Count; i++)
                        WriteOverride(w, "/word/" + footers[i].FileName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
                    if (hasFootnotes)
                        WriteOverride(w, "/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml");
                    if (hasEndnotes)
                        WriteOverride(w, "/word/endnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml");
                    if (hasComments)
                        WriteOverride(w, "/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml");
                    WriteOverride(w, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
                    WriteOverride(w, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
                    if (doc != null && doc.PreservedContentTypes != null)
                    {
                        for (int i = 0; i < doc.PreservedContentTypes.Count; i++)
                        {
                            PreservedContentType preserved = doc.PreservedContentTypes[i];
                            if (string.IsNullOrEmpty(preserved.PartName)) continue;
                            WriteOverride(w, "/" + preserved.PartName.TrimStart('/'), preserved.ContentType);
                        }
                    }
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static void WriteDefault(XmlWriter w, string ext, string type)
        {
            w.WriteStartElement("Default", CT.NamespaceName);
            w.WriteAttributeString("Extension", ext);
            w.WriteAttributeString("ContentType", type);
            w.WriteEndElement();
        }

        private static void WriteOverride(XmlWriter w, string part, string type)
        {
            w.WriteStartElement("Override", CT.NamespaceName);
            w.WriteAttributeString("PartName", part);
            w.WriteAttributeString("ContentType", type);
            w.WriteEndElement();
        }

        private static byte[] BuildRootRels()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("Relationships", REL.NamespaceName);
                    WriteRel(w, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "word/document.xml", null);
                    WriteRel(w, "rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml", null);
                    WriteRel(w, "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml", null);
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildSettingsXml(DocumentModel doc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("w", "settings", W.NamespaceName);
                    bool evenOdd = false;
                    for (int i = 0; i < doc.Sections.Count; i++)
                        if (doc.Sections[i].DifferentOddEven) evenOdd = true;
                    if (evenOdd)
                        w.WriteElementString("w", "evenAndOddHeaders", W.NamespaceName, string.Empty);
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildDocumentRels(DocumentModel doc, List<ImagePart> images, bool hasBullets,
            List<HeaderFooterPart> headers, List<HeaderFooterPart> footers,
            bool hasFootnotes, bool hasEndnotes, bool hasComments, List<EmbeddedPart> embedded, List<HyperlinkPart> hyperlinks)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("Relationships", REL.NamespaceName);
                    WriteRel(w, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml", null);
                    WriteRel(w, "rIdSettings", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml", null);
                    if (hasBullets)
                        WriteRel(w, "rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml", null);
                    for (int i = 0; i < headers.Count; i++)
                        WriteRel(w, headers[i].RelId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                            headers[i].FileName, null);
                    for (int i = 0; i < footers.Count; i++)
                        WriteRel(w, footers[i].RelId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                            footers[i].FileName, null);
                    if (hasFootnotes)
                        WriteRel(w, "rIdFtn", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", "footnotes.xml", null);
                    if (hasEndnotes)
                        WriteRel(w, "rIdEnd", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes", "endnotes.xml", null);
                    if (hasComments)
                        WriteRel(w, "rIdCmt", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml", null);
                    for (int i = 0; i < embedded.Count; i++)
                        WriteRel(w, embedded[i].RelId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                            "embeddings/" + embedded[i].FileName, null);
                    if (hyperlinks != null)
                    {
                        for (int i = 0; i < hyperlinks.Count; i++)
                            WriteRel(w, hyperlinks[i].RelId,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                                hyperlinks[i].Url, "External");
                    }
                    if (doc != null && doc.PreservedRelationships != null)
                    {
                        for (int i = 0; i < doc.PreservedRelationships.Count; i++)
                        {
                            PreservedRelationship preserved = doc.PreservedRelationships[i];
                            string mode = string.IsNullOrEmpty(preserved.TargetMode) ? null : preserved.TargetMode;
                            WriteRel(w, "rIdPreserved" + (i + 1), preserved.Type, preserved.Target, mode);
                        }
                    }
                    for (int i = 0; i < images.Count; i++)
                    {
                        WriteRel(w, images[i].RelId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                            "media/" + images[i].FileName, null);
                    }
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static void WriteRel(XmlWriter w, string id, string type, string target, string mode)
        {
            w.WriteStartElement("Relationship", REL.NamespaceName);
            w.WriteAttributeString("Id", id);
            w.WriteAttributeString("Type", type);
            w.WriteAttributeString("Target", target);
            if (mode != null) w.WriteAttributeString("TargetMode", mode);
            w.WriteEndElement();
        }

        private static byte[] BuildDocumentXml(DocumentModel doc, List<ImagePart> images,
            List<HeaderFooterPart> headers, List<HeaderFooterPart> footers,
            List<EmbeddedPart> embedded, List<HyperlinkPart> hyperlinks)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("w", "document", W.NamespaceName);
                    w.WriteAttributeString("xmlns", "r", null, R.NamespaceName);
                    w.WriteAttributeString("xmlns", "wp", null, WP.NamespaceName);
                    w.WriteAttributeString("xmlns", "a", null, A.NamespaceName);
                    w.WriteAttributeString("xmlns", "pic", null, "http://schemas.openxmlformats.org/drawingml/2006/picture");
                    w.WriteAttributeString("xmlns", "v", null, VML.NamespaceName);
                    w.WriteAttributeString("xmlns", "o", null, OLE.NamespaceName);
                    w.WriteStartElement("w", "body", W.NamespaceName);

                    doc.EnsureSections();
                    int imageCounter = 0;
                    for (int si = 0; si < doc.Sections.Count; si++)
                    {
                        DocumentSection section = doc.Sections[si];
                        int start = Math.Max(0, Math.Min(doc.Paragraphs.Count, section.StartParagraph));
                        int end = si + 1 < doc.Sections.Count
                            ? Math.Max(start, doc.Sections[si + 1].StartParagraph - 1)
                            : doc.Paragraphs.Count - 1;
                        for (int pi = start; pi <= end && pi < doc.Paragraphs.Count; pi++)
                        {
                            Paragraph paragraph = doc.Paragraphs[pi];
                            bool close = si + 1 < doc.Sections.Count && pi == end;
                            if (paragraph.TableRows != null)
                            {
                                WriteTable(w, paragraph, hyperlinks);
                                if (close)
                                {
                                    w.WriteStartElement("w", "p", W.NamespaceName);
                                    w.WriteStartElement("w", "pPr", W.NamespaceName);
                                    WriteSectionElement(w, section, headers, footers);
                                    w.WriteEndElement();
                                    w.WriteEndElement();
                                }
                            }
                            else
                            {
                                WriteParagraph(w, paragraph, images, ref imageCounter,
                                    close ? section : null, headers, footers, embedded, hyperlinks);
                            }
                        }
                    }
                    if (doc.Paragraphs.Count == 0)
                    {
                        w.WriteStartElement("w", "p", W.NamespaceName);
                        w.WriteEndElement();
                    }

                    WriteSectionElement(w, doc.Sections[doc.Sections.Count - 1], headers, footers);

                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildHeaderFooterXml(HeaderFooterPart part, List<ImagePart> images)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("w", part.IsHeader ? "hdr" : "ftr", W.NamespaceName);
                    w.WriteAttributeString("xmlns", "r", null, R.NamespaceName);
                    w.WriteAttributeString("xmlns", "wp", null, WP.NamespaceName);
                    w.WriteAttributeString("xmlns", "a", null, A.NamespaceName);
                    w.WriteAttributeString("xmlns", "pic", null, "http://schemas.openxmlformats.org/drawingml/2006/picture");
                    string normalized = (part.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
                    string[] lines = normalized.Split('\n');
                    for (int i = 0; i < lines.Length; i++)
                    {
                        w.WriteStartElement("w", "p", W.NamespaceName);
                        WriteHeaderFooterAlignment(w, part.Alignment);
                        WriteHeaderFooterRuns(w, lines[i], part.Format);
                        w.WriteEndElement();
                    }
                    int counter = 0;
                    for (int i = 0; part.Images != null && i < part.Images.Count; i++)
                    {
                        w.WriteStartElement("w", "p", W.NamespaceName);
                        WriteImageRun(w, part.Images[i], images, ref counter);
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildHeaderFooterRels(HeaderFooterPart part, List<ImagePart> images)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("Relationships", REL.NamespaceName);
                    HashSet<string> written = new HashSet<string>(StringComparer.Ordinal);
                    for (int i = 0; part.Images != null && i < part.Images.Count; i++)
                    {
                        ImageInline image = part.Images[i];
                        for (int j = 0; j < images.Count; j++)
                        {
                            if (images[j].Data == image.Data
                                || (images[j].Data.Length == image.Data.Length && SameBytes(images[j].Data, image.Data)))
                            {
                                if (written.Add(images[j].RelId))
                                    WriteRel(w, images[j].RelId,
                                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                                        "media/" + images[j].FileName, null);
                                break;
                            }
                        }
                    }
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static void WriteHeaderFooterAlignment(XmlWriter w, ParagraphAlignment alignment)
        {
            w.WriteStartElement("w", "pPr", W.NamespaceName);
            w.WriteStartElement("w", "jc", W.NamespaceName);
            string value = alignment == ParagraphAlignment.Center ? "center"
                : alignment == ParagraphAlignment.Right ? "right"
                : alignment == ParagraphAlignment.Justify ? "both" : "left";
            w.WriteAttributeString("w", "val", W.NamespaceName, value);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteHeaderFooterRuns(XmlWriter w, string line, RunFormat format)
        {
            int pos = 0;
            while (pos < line.Length)
            {
                int page = line.IndexOf("{PAGE}", pos, StringComparison.Ordinal);
                int total = line.IndexOf("{NUMPAGES}", pos, StringComparison.Ordinal);
                int next = page;
                string field = "PAGE";
                if (next < 0 || (total >= 0 && total < next)) { next = total; field = "NUMPAGES"; }
                if (next < 0)
                {
                    WriteHeaderFooterText(w, line.Substring(pos), format);
                    break;
                }
                if (next > pos) WriteHeaderFooterText(w, line.Substring(pos, next - pos), format);
                WriteHeaderFooterField(w, field, format);
                pos = next + (field == "PAGE" ? 6 : 10);
            }
            if (line.Length == 0) WriteHeaderFooterText(w, string.Empty, format);
        }

        private static void WriteHeaderFooterText(XmlWriter w, string text, RunFormat format)
        {
            if (text.Length == 0) return;
            w.WriteStartElement("w", "r", W.NamespaceName);
            WriteHeaderFooterRunProperties(w, format);
            w.WriteStartElement("w", "t", W.NamespaceName);
            w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(text);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteHeaderFooterRunProperties(XmlWriter w, RunFormat format)
        {
            RunFormat f = format ?? new RunFormat("Calibri", 18);
            w.WriteStartElement("w", "rPr", W.NamespaceName);
            w.WriteStartElement("w", "rFonts", W.NamespaceName);
            w.WriteAttributeString("w", "ascii", W.NamespaceName, f.FontFamily);
            w.WriteAttributeString("w", "hAnsi", W.NamespaceName, f.FontFamily);
            w.WriteEndElement();
            w.WriteStartElement("w", "sz", W.NamespaceName);
            w.WriteAttributeString("w", "val", W.NamespaceName, Math.Max(2, f.SizeHalfPoints).ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            if (f.Bold) w.WriteElementString("w", "b", W.NamespaceName, string.Empty);
            if (f.Italic) w.WriteElementString("w", "i", W.NamespaceName, string.Empty);
            if (f.Underline) w.WriteElementString("w", "u", W.NamespaceName, string.Empty);
            if (f.ForeColor.HasValue)
            {
                w.WriteStartElement("w", "color", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName,
                    f.ForeColor.Value.R.ToString("X2") + f.ForeColor.Value.G.ToString("X2") + f.ForeColor.Value.B.ToString("X2"));
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        private static void WriteHeaderFooterField(XmlWriter w, string field, RunFormat format)
        {
            w.WriteStartElement("w", "r", W.NamespaceName);
            WriteHeaderFooterRunProperties(w, format);
            w.WriteStartElement("w", "fldChar", W.NamespaceName);
            w.WriteAttributeString("w", "fldCharType", W.NamespaceName, "begin");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("w", "r", W.NamespaceName);
            WriteHeaderFooterRunProperties(w, format);
            w.WriteStartElement("w", "instrText", W.NamespaceName);
            w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(" " + field + " ");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("w", "r", W.NamespaceName);
            WriteHeaderFooterRunProperties(w, format);
            w.WriteStartElement("w", "fldChar", W.NamespaceName);
            w.WriteAttributeString("w", "fldCharType", W.NamespaceName, "separate");
            w.WriteEndElement();
            w.WriteEndElement();

            WriteHeaderFooterText(w, "1", format);

            w.WriteStartElement("w", "r", W.NamespaceName);
            WriteHeaderFooterRunProperties(w, format);
            w.WriteStartElement("w", "fldChar", W.NamespaceName);
            w.WriteAttributeString("w", "fldCharType", W.NamespaceName, "end");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteTable(XmlWriter w, Paragraph p, List<HyperlinkPart> hyperlinks)
        {
            p.EnsureTableCells();
            if (p.TableRows == null && p.TableCells != null) p.SyncTableRows();
            string[][] rows = p.TableRows;
            if (rows == null) return;
            int columns = 0;
            for (int r = 0; r < rows.Length; r++)
                if (rows[r] != null && rows[r].Length > columns) columns = rows[r].Length;
            if (columns == 0) return;

            int width = p.TableWidthTwips > 0 ? p.TableWidthTwips : 9000;
            int cellWidth = Math.Max(300, width / columns);

            w.WriteStartElement("w", "tbl", W.NamespaceName);
            w.WriteStartElement("w", "tblPr", W.NamespaceName);
            if (!string.IsNullOrEmpty(p.TableStyleName))
            {
                w.WriteStartElement("w", "tblStyle", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, p.TableStyleName);
                w.WriteEndElement();
            }
            w.WriteStartElement("w", "tblW", W.NamespaceName);
            w.WriteAttributeString("w", "w", W.NamespaceName, width.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "type", W.NamespaceName, "dxa");
            w.WriteEndElement();
            w.WriteStartElement("w", "tblBorders", W.NamespaceName);
            string[] borderNames = new string[] { "top", "left", "bottom", "right", "insideH", "insideV" };
            for (int i = 0; i < borderNames.Length; i++)
            {
                w.WriteStartElement("w", borderNames[i], W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, p.TableBorderColor.HasValue ? "single" : "nil");
                if (p.TableBorderColor.HasValue)
                {
                    w.WriteAttributeString("w", "sz", W.NamespaceName, Math.Max(2, p.TableBorderSize * 8).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("w", "color", W.NamespaceName, ColorHex(p.TableBorderColor.Value));
                }
                w.WriteEndElement();
            }
            w.WriteEndElement();
            if (p.TableShading.HasValue)
            {
                w.WriteStartElement("w", "shd", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, "clear");
                w.WriteAttributeString("w", "color", W.NamespaceName, "auto");
                w.WriteAttributeString("w", "fill", W.NamespaceName, ColorHex(p.TableShading.Value));
                w.WriteEndElement();
            }
            w.WriteEndElement();

            w.WriteStartElement("w", "tblGrid", W.NamespaceName);
            for (int c = 0; c < columns; c++)
            {
                w.WriteStartElement("w", "gridCol", W.NamespaceName);
                w.WriteAttributeString("w", "w", W.NamespaceName, cellWidth.ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
            }
            w.WriteEndElement();

            p.EnsureTableCells();
            TableCell[][] tableCells = p.TableCells;
            for (int r = 0; r < rows.Length; r++)
            {
                w.WriteStartElement("w", "tr", W.NamespaceName);
                if (r == 0 && p.RepeatTableHeader)
                {
                    w.WriteStartElement("w", "trPr", W.NamespaceName);
                    w.WriteStartElement("w", "tblHeader", W.NamespaceName);
                    w.WriteEndElement();
                    w.WriteEndElement();
                }
                for (int c = 0; c < columns; )
                {
                    TableCell owner = FindOwnerCell(tableCells, r, c);
                    if (owner == null)
                    {
                        WriteTableCell(w, string.Empty, 1, 1, false, cellWidth, null, hyperlinks);
                        c++;
                        continue;
                    }
                    int ownerRow, ownerColumn;
                    FindOwnerPosition(tableCells, owner, out ownerRow, out ownerColumn);
                    if (ownerRow == r && ownerColumn == c)
                    {
                        WriteTableCell(w, owner.Text, Math.Max(1, owner.ColSpan), Math.Max(1, owner.RowSpan), false, cellWidth, owner, hyperlinks);
                        c += Math.Max(1, owner.ColSpan);
                    }
                    else if (ownerRow < r)
                    {
                        if (ownerColumn == c) WriteTableCell(w, string.Empty, 1, 1, true, cellWidth, owner, hyperlinks);
                        c++;
                    }
                    else if (ownerRow == r && ownerColumn < c)
                    {
                        c++;
                    }
                    else
                    {
                        WriteTableCell(w, string.Empty, 1, 1, false, cellWidth, null, hyperlinks);
                        c++;
                    }
                }
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        private static void WriteTableCell(XmlWriter w, string text, int colSpan, int rowSpan,
            bool verticalContinue, int cellWidth, TableCell cell, List<HyperlinkPart> hyperlinks)
        {
            w.WriteStartElement("w", "tc", W.NamespaceName);
            w.WriteStartElement("w", "tcPr", W.NamespaceName);
            w.WriteStartElement("w", "tcW", W.NamespaceName);
            w.WriteAttributeString("w", "w", W.NamespaceName, (cellWidth * colSpan).ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "type", W.NamespaceName, "dxa");
            w.WriteEndElement();
            if (colSpan > 1)
            {
                w.WriteStartElement("w", "gridSpan", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, colSpan.ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
            }
            if (cell != null && cell.Shading.HasValue)
            {
                w.WriteStartElement("w", "shd", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, "clear");
                w.WriteAttributeString("w", "color", W.NamespaceName, "auto");
                w.WriteAttributeString("w", "fill", W.NamespaceName, ColorHex(cell.Shading.Value));
                w.WriteEndElement();
            }
            if (cell != null && cell.BorderColor.HasValue)
            {
                w.WriteStartElement("w", "tcBorders", W.NamespaceName);
                string[] sides = new string[] { "top", "left", "bottom", "right" };
                for (int i = 0; i < sides.Length; i++)
                {
                    w.WriteStartElement("w", sides[i], W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, "single");
                    w.WriteAttributeString("w", "sz", W.NamespaceName, Math.Max(2, cell.BorderSize * 8).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("w", "color", W.NamespaceName, ColorHex(cell.BorderColor.Value));
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }
            if (rowSpan > 1)
            {
                w.WriteStartElement("w", "vMerge", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, "restart");
                w.WriteEndElement();
            }
            else if (verticalContinue)
            {
                w.WriteStartElement("w", "vMerge", W.NamespaceName);
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteStartElement("w", "p", W.NamespaceName);
            if (!verticalContinue && cell != null && cell.Runs.Count > 0)
            {
                for (int ri = 0; ri < cell.Runs.Count; ri++)
                {
                    TableCellRun cellRun = cell.Runs[ri];
                    if (cellRun == null || string.IsNullOrEmpty(cellRun.Text)) continue;
                    string runLinkId = !string.IsNullOrEmpty(cellRun.LinkUrl)
                        ? HyperlinkRelId(hyperlinks, cellRun.LinkUrl) : null;
                    if (!string.IsNullOrEmpty(runLinkId))
                    {
                        w.WriteStartElement("w", "hyperlink", W.NamespaceName);
                        w.WriteAttributeString("r", "id", R.NamespaceName, runLinkId);
                    }
                    TextInline fieldRun = new TextInline(cellRun.Text,
                        cellRun.Format == null ? new RunFormat("Calibri", 21) : cellRun.Format.Clone());
                    fieldRun.Format.LinkUrl = cellRun.LinkUrl;
                    WriteTextInline(w, fieldRun);
                    if (!string.IsNullOrEmpty(runLinkId)) w.WriteEndElement();
                }
            }
            else
            {
                string cellLinkId = cell != null ? HyperlinkRelId(hyperlinks, cell.LinkUrl) : null;
                if (!string.IsNullOrEmpty(cellLinkId))
                {
                    w.WriteStartElement("w", "hyperlink", W.NamespaceName);
                    w.WriteAttributeString("r", "id", R.NamespaceName, cellLinkId);
                }
                if (!string.IsNullOrEmpty(text))
                {
                    TextInline fieldRun = new TextInline(text,
                        cell != null && cell.Format != null ? cell.Format.Clone() : new RunFormat("Calibri", 21)
                        {
                            Bold = cell != null && cell.Bold,
                            Italic = cell != null && cell.Italic
                        });
                    WriteTextInline(w, fieldRun);
                }
                if (!string.IsNullOrEmpty(cellLinkId)) w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static string HyperlinkRelId(List<HyperlinkPart> hyperlinks, string url)
        {
            if (hyperlinks == null || string.IsNullOrEmpty(url)) return null;
            for (int i = 0; i < hyperlinks.Count; i++)
                if (string.Equals(hyperlinks[i].Url, url, StringComparison.Ordinal))
                    return hyperlinks[i].RelId;
            return null;
        }

        private static TableCell FindOwnerCell(TableCell[][] cells, int row, int column)
        {
            if (cells == null) return null;
            for (int r = 0; r < cells.Length; r++)
            {
                if (cells[r] == null) continue;
                for (int c = 0; c < cells[r].Length; c++)
                {
                    TableCell cell = cells[r][c];
                    if (cell == null || cell.Covered) continue;
                    if (row >= r && row < r + Math.Max(1, cell.RowSpan) &&
                        column >= c && column < c + Math.Max(1, cell.ColSpan)) return cell;
                }
            }
            return null;
        }

        private static void FindOwnerPosition(TableCell[][] cells, TableCell owner, out int row, out int column)
        {
            row = 0;
            column = 0;
            for (int r = 0; r < cells.Length; r++)
            {
                if (cells[r] == null) continue;
                for (int c = 0; c < cells[r].Length; c++)
                {
                    if (cells[r][c] == owner) { row = r; column = c; return; }
                }
            }
        }

        private static string ParagraphFormatKey(Paragraph p)
        {
            StringBuilder key = new StringBuilder();
            key.Append((int)p.Alignment).Append('|').Append(p.LeftIndentTwips).Append('|')
                .Append(p.RightIndentTwips).Append('|').Append(p.FirstLineIndentTwips).Append('|')
                .Append(p.SpaceBeforeTwips).Append('|').Append(p.SpaceAfterTwips).Append('|')
                .Append(p.LineSpacing).Append('|').Append(p.Bullet ? 1 : 0).Append('|')
                .Append(p.BulletLevel).Append('|').Append(p.Numbered ? 1 : 0).Append('|')
                .Append(p.NumberLevel).Append('|').Append(p.PageBreakBefore ? 1 : 0)
                .Append('|').Append(p.StyleName).Append('|').Append(p.RepeatTableHeader ? 1 : 0);
            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline text = p.Inlines[i] as TextInline;
                if (text == null) continue;
                RunFormat f = text.Format ?? new RunFormat();
                key.Append('#').Append(f.FontFamily).Append(',').Append(f.SizeHalfPoints).Append(',')
                    .Append(f.Bold ? 1 : 0).Append(f.Italic ? 1 : 0).Append(f.Underline ? 1 : 0)
                    .Append(f.Strikeout ? 1 : 0).Append(',').Append(f.ForeColor.HasValue ? f.ForeColor.Value.ToArgb() : 0)
                    .Append(',').Append(f.BackColor.HasValue ? f.BackColor.Value.ToArgb() : 0)
                    .Append(',').Append(f.Script).Append(',').Append(f.LinkUrl);
            }
            return key.ToString();
        }

        private static string ModelParagraphText(Paragraph p)
        {
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline inline = p.Inlines[i] as TextInline;
                if (inline != null) text.Append(inline.Text);
            }
            return text.ToString();
        }

        private static void WriteParagraph(XmlWriter w, Paragraph p, List<ImagePart> images, ref int imageCounter)
        {
            WriteParagraph(w, p, images, ref imageCounter, null, null, null, null, null);
        }

        private static void WriteParagraph(XmlWriter w, Paragraph p, List<ImagePart> images, ref int imageCounter,
            DocumentSection closeSection, List<HeaderFooterPart> headers, List<HeaderFooterPart> footers,
            List<EmbeddedPart> embedded, List<HyperlinkPart> hyperlinks)
        {
            if (closeSection == null && p.PreservedXml != null
                && string.Equals(ModelParagraphText(p), p.PreservedText, StringComparison.Ordinal)
                && string.Equals(ParagraphFormatKey(p), p.PreservedFormatKey, StringComparison.Ordinal))
            {
                w.WriteRaw(p.PreservedXml);
                return;
            }
            w.WriteStartElement("w", "p", W.NamespaceName);
            if (p.PageBreakBefore)
            {
                w.WriteStartElement("w", "r", W.NamespaceName);
                w.WriteStartElement("w", "br", W.NamespaceName);
                w.WriteAttributeString("w", "type", W.NamespaceName, "page");
                w.WriteEndElement();
                w.WriteEndElement();
            }

            bool hasPPr = p.Alignment != ParagraphAlignment.Left
                || p.LeftIndentTwips != 0 || p.RightIndentTwips != 0 || p.FirstLineIndentTwips != 0
                || p.SpaceBeforeTwips != 0 || p.SpaceAfterTwips != 0 || p.LineSpacing != 0
                || (p.TabStopsTwips != null && p.TabStopsTwips.Length > 0)
                || p.Bullet || p.Numbered || p.KeepLinesTogether
                || p.KeepWithNext || !p.WidowControl || p.OutlineLevel != 0
                || p.NoSpaceBetweenSameStyle || !string.IsNullOrEmpty(p.StyleName) || closeSection != null;

            if (hasPPr)
            {
                w.WriteStartElement("w", "pPr", W.NamespaceName);
                if (!string.IsNullOrEmpty(p.StyleName))
                {
                    w.WriteStartElement("w", "pStyle", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, DocumentStyles.StyleId(p.StyleName));
                    w.WriteEndElement();
                }
                if (p.Bullet || p.Numbered)
                {
                    w.WriteStartElement("w", "numPr", W.NamespaceName);
                    w.WriteStartElement("w", "ilvl", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName,
                        Math.Max(0, Math.Min(8, p.Numbered ? p.NumberLevel : p.BulletLevel)).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                    w.WriteStartElement("w", "numId", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, p.Numbered ? "2" : "1");
                    w.WriteEndElement();
                    w.WriteEndElement();
                }
                if (p.TabStopsTwips != null && p.TabStopsTwips.Length > 0)
                {
                    w.WriteStartElement("w", "tabs", W.NamespaceName);
                    for (int i = 0; i < p.TabStopsTwips.Length; i++)
                    {
                        w.WriteStartElement("w", "tab", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, "left");
                        w.WriteAttributeString("w", "pos", W.NamespaceName, p.TabStopsTwips[i].ToString(CultureInfo.InvariantCulture));
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                }
                if (p.SpaceBeforeTwips != 0 || p.SpaceAfterTwips != 0 || p.LineSpacing != 0)
                {
                    w.WriteStartElement("w", "spacing", W.NamespaceName);
                    if (p.SpaceBeforeTwips != 0)
                        w.WriteAttributeString("w", "before", W.NamespaceName, p.SpaceBeforeTwips.ToString(CultureInfo.InvariantCulture));
                    if (p.SpaceAfterTwips != 0)
                        w.WriteAttributeString("w", "after", W.NamespaceName, p.SpaceAfterTwips.ToString(CultureInfo.InvariantCulture));
                    if (p.LineSpacing == 1)
                    {
                        w.WriteAttributeString("w", "line", W.NamespaceName, "360");
                        w.WriteAttributeString("w", "lineRule", W.NamespaceName, "auto");
                    }
                    else if (p.LineSpacing == 2)
                    {
                        w.WriteAttributeString("w", "line", W.NamespaceName, "480");
                        w.WriteAttributeString("w", "lineRule", W.NamespaceName, "auto");
                    }
                    w.WriteEndElement();
                }
                if (p.LeftIndentTwips != 0 || p.RightIndentTwips != 0 || p.FirstLineIndentTwips != 0)
                {
                    w.WriteStartElement("w", "ind", W.NamespaceName);
                    if (p.LeftIndentTwips != 0)
                        w.WriteAttributeString("w", "left", W.NamespaceName, p.LeftIndentTwips.ToString(CultureInfo.InvariantCulture));
                    if (p.RightIndentTwips != 0)
                        w.WriteAttributeString("w", "right", W.NamespaceName, p.RightIndentTwips.ToString(CultureInfo.InvariantCulture));
                    if (p.FirstLineIndentTwips > 0)
                        w.WriteAttributeString("w", "firstLine", W.NamespaceName, p.FirstLineIndentTwips.ToString(CultureInfo.InvariantCulture));
                    else if (p.FirstLineIndentTwips < 0)
                        w.WriteAttributeString("w", "hanging", W.NamespaceName, (-p.FirstLineIndentTwips).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                }
                if (p.Alignment != ParagraphAlignment.Left)
                {
                    w.WriteStartElement("w", "jc", W.NamespaceName);
                    string val = "left";
                    if (p.Alignment == ParagraphAlignment.Center) val = "center";
                    else if (p.Alignment == ParagraphAlignment.Right) val = "right";
                    else if (p.Alignment == ParagraphAlignment.Justify) val = "both";
                    w.WriteAttributeString("w", "val", W.NamespaceName, val);
                    w.WriteEndElement();
                }
                if (p.KeepLinesTogether) w.WriteElementString("w", "keepLines", W.NamespaceName, string.Empty);
                if (p.KeepWithNext) w.WriteElementString("w", "keepNext", W.NamespaceName, string.Empty);
                if (!p.WidowControl)
                {
                    w.WriteStartElement("w", "widowControl", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, "0");
                    w.WriteEndElement();
                }
                if (p.NoSpaceBetweenSameStyle) w.WriteElementString("w", "contextualSpacing", W.NamespaceName, string.Empty);
                if (p.OutlineLevel != 0)
                {
                    w.WriteStartElement("w", "outlineLvl", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, Math.Max(0, Math.Min(8, p.OutlineLevel)).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                }
                if (closeSection != null) WriteSectionElement(w, closeSection, headers, footers);
                w.WriteEndElement();
            }

            for (int i = 0; i < p.Inlines.Count; i++)
            {
                TextInline text = p.Inlines[i] as TextInline;
                if (text != null)
                {
                    HyperlinkPart link = null;
                    if (hyperlinks != null && !string.IsNullOrEmpty(text.Format.LinkUrl))
                        for (int h = 0; h < hyperlinks.Count; h++)
                            if (string.Equals(hyperlinks[h].Url, text.Format.LinkUrl, StringComparison.Ordinal))
                            {
                                link = hyperlinks[h];
                                break;
                            }
                    if (link != null)
                    {
                        w.WriteStartElement("w", "hyperlink", W.NamespaceName);
                        w.WriteAttributeString("r", "id", R.NamespaceName, link.RelId);
                        WriteTextInline(w, text);
                        w.WriteEndElement();
                    }
                    else
                    {
                        WriteTextInline(w, text);
                    }
                    continue;
                }
                NoteInline note = p.Inlines[i] as NoteInline;
                if (note != null && note.Note != null)
                {
                    w.WriteStartElement("w", "r", W.NamespaceName);
                    w.WriteStartElement("w", note.Note.Endnote ? "endnoteReference" : "footnoteReference", W.NamespaceName);
                    w.WriteAttributeString("w", "id", W.NamespaceName, note.Note.Id.ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                    w.WriteEndElement();
                    continue;
                }
                CommentInline comment = p.Inlines[i] as CommentInline;
                if (comment != null && comment.Comment != null)
                {
                    w.WriteStartElement("w", "r", W.NamespaceName);
                    w.WriteStartElement("w", "commentReference", W.NamespaceName);
                    w.WriteAttributeString("w", "id", W.NamespaceName, comment.Comment.Id.ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                    w.WriteEndElement();
                    continue;
                }
                ImageInline img = p.Inlines[i] as ImageInline;
                if (img != null)
                {
                    if (WriteImageRun(w, img, images, ref imageCounter))
                        continue;
                }
            }

            for (int i = 0; i < p.FloatingObjects.Count; i++)
            {
                FloatingObject floating = p.FloatingObjects[i];
                if (floating.Image != null && floating.Kind == FloatingKind.Image)
                    WriteFloatingImageRun(w, floating, images, ref imageCounter);
                else if (floating.Kind == FloatingKind.TextBox || floating.Kind == FloatingKind.Shape)
                    // VML is understood by Word 2007, Word 2010+ and
                    // LibreOffice; the wps DrawingML vocabulary used before
                    // is a Word 2010 extension and can make Word 2007 report
                    // an unreadable document.
                    WriteFloatingVmlRun(w, floating);
                else if (floating.Kind == FloatingKind.Embedded && floating.Image != null)
                    WriteFloatingVmlImageRun(w, floating, images, ref imageCounter, embedded);
            }

            if (p.Inlines.Count == 0 && p.FloatingObjects.Count == 0)
            {
                w.WriteStartElement("w", "r", W.NamespaceName);
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        private static void WriteHeaderFooterReference(XmlWriter w, string element,
            HeaderFooterPart part)
        {
            if (part == null) return;
            w.WriteStartElement("w", element, W.NamespaceName);
            w.WriteAttributeString("w", "type", W.NamespaceName, part.Type);
            w.WriteAttributeString("r", "id", R.NamespaceName, part.RelId);
            w.WriteEndElement();
        }

        private static void WriteSectionElement(XmlWriter w, DocumentSection section,
            List<HeaderFooterPart> headers, List<HeaderFooterPart> footers)
        {
            if (section == null) return;
            w.WriteStartElement("w", "sectPr", W.NamespaceName);
            WriteHeaderFooterReference(w, "headerReference", FindHeaderFooter(headers, section, "default"));
            WriteHeaderFooterReference(w, "headerReference", FindHeaderFooter(headers, section, "first"));
            WriteHeaderFooterReference(w, "headerReference", FindHeaderFooter(headers, section, "even"));
            WriteHeaderFooterReference(w, "footerReference", FindHeaderFooter(footers, section, "default"));
            WriteHeaderFooterReference(w, "footerReference", FindHeaderFooter(footers, section, "first"));
            WriteHeaderFooterReference(w, "footerReference", FindHeaderFooter(footers, section, "even"));
            w.WriteStartElement("w", "pgSz", W.NamespaceName);
            int pw = section.PageWidthTwips;
            int ph = section.PageHeightTwips;
            w.WriteAttributeString("w", "w", W.NamespaceName, pw.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "h", W.NamespaceName, ph.ToString(CultureInfo.InvariantCulture));
            if (pw > ph) w.WriteAttributeString("w", "orient", W.NamespaceName, "landscape");
            w.WriteEndElement();
            w.WriteStartElement("w", "pgMar", W.NamespaceName);
            w.WriteAttributeString("w", "top", W.NamespaceName, section.MarginTopTwips.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "right", W.NamespaceName, section.MarginRightTwips.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "bottom", W.NamespaceName, section.MarginBottomTwips.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "left", W.NamespaceName, section.MarginLeftTwips.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "header", W.NamespaceName, Math.Max(0, section.HeaderDistanceTwips).ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "footer", W.NamespaceName, Math.Max(0, section.FooterDistanceTwips).ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("w", "gutter", W.NamespaceName, "0");
            w.WriteEndElement();
            if (section.ColumnCount > 1)
            {
                w.WriteStartElement("w", "cols", W.NamespaceName);
                w.WriteAttributeString("w", "num", W.NamespaceName, section.ColumnCount.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("w", "space", W.NamespaceName, section.ColumnSpacingTwips.ToString(CultureInfo.InvariantCulture));
                if (section.ColumnsSeparator) w.WriteAttributeString("w", "sep", W.NamespaceName, "1");
                w.WriteEndElement();
            }
            if (section.BreakType != SectionBreakType.NextPage)
            {
                w.WriteStartElement("w", "type", W.NamespaceName);
                string value = section.BreakType == SectionBreakType.Continuous ? "continuous"
                    : section.BreakType == SectionBreakType.EvenPage ? "evenPage" : "oddPage";
                w.WriteAttributeString("w", "val", W.NamespaceName, value);
                w.WriteEndElement();
            }
            if (section.TitlePage) w.WriteStartElement("w", "titlePg", W.NamespaceName);
            if (section.TitlePage) w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteTextInline(XmlWriter w, TextInline text)
        {
            bool field = text.Field != FieldKind.None;
            if (field)
            {
                w.WriteStartElement("w", "fldSimple", W.NamespaceName);
                w.WriteAttributeString("w", "instr", W.NamespaceName,
                    string.IsNullOrEmpty(text.FieldInstruction) ? FieldText.Instruction(text.Field, null) : text.FieldInstruction);
            }
            if (text.Revision == RevisionKind.Insert)
            {
                w.WriteStartElement("w", "ins", W.NamespaceName);
                w.WriteAttributeString("w", "id", W.NamespaceName, NextRevisionId().ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("w", "author", W.NamespaceName, text.RevisionAuthor ?? "Mini Word");
                w.WriteAttributeString("w", "date", W.NamespaceName, RevisionDate(text.RevisionDate));
                WriteTextRun(w, text, false);
                w.WriteEndElement();
            }
            else if (text.Revision == RevisionKind.Delete)
            {
                w.WriteStartElement("w", "del", W.NamespaceName);
                w.WriteAttributeString("w", "id", W.NamespaceName, NextRevisionId().ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("w", "author", W.NamespaceName, text.RevisionAuthor ?? "Mini Word");
                w.WriteAttributeString("w", "date", W.NamespaceName, RevisionDate(text.RevisionDate));
                WriteTextRun(w, text, true);
                w.WriteEndElement();
            }
            else
            {
                WriteTextRun(w, text, false);
            }
            if (field) w.WriteEndElement();
        }

        private static int _revisionId;

        private static int NextRevisionId()
        {
            return System.Threading.Interlocked.Increment(ref _revisionId);
        }

        private static string RevisionDate(string value)
        {
            DateTime date;
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date))
                return date.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
            return DateTime.UtcNow.ToString("s", CultureInfo.InvariantCulture) + "Z";
        }
        private static void WriteTextRun(XmlWriter w, TextInline text)
        {
            WriteTextRun(w, text, false);
        }

        private static void WriteTextRun(XmlWriter w, TextInline text, bool deleted)
        {
            w.WriteStartElement("w", "r", W.NamespaceName);
            WriteRunProperties(w, text.Format);
            string value = text.Text ?? string.Empty;
            StringBuilder segment = new StringBuilder();
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c == '\n')
                {
                    FlushText(w, segment, deleted);
                    w.WriteStartElement("w", "br", W.NamespaceName);
                    w.WriteEndElement();
                }
                else if (c == '\t')
                {
                    FlushText(w, segment, deleted);
                    w.WriteStartElement("w", "tab", W.NamespaceName);
                    w.WriteEndElement();
                }
                else if (c == '\r')
                {
                    // ignore
                }
                else
                {
                    segment.Append(c);
                }
            }
            FlushText(w, segment, deleted);
            w.WriteEndElement();
        }

        private static void FlushText(XmlWriter w, StringBuilder segment, bool deleted)
        {
            if (segment.Length == 0) return;
            w.WriteStartElement("w", "t", W.NamespaceName);
            w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(segment.ToString());
            w.WriteEndElement();
            segment.Length = 0;
        }

        private static void WriteRunProperties(XmlWriter w, RunFormat f)
        {
            w.WriteStartElement("w", "rPr", W.NamespaceName);
            if (!string.IsNullOrEmpty(f.FontFamily))
            {
                w.WriteStartElement("w", "rFonts", W.NamespaceName);
                w.WriteAttributeString("w", "ascii", W.NamespaceName, f.FontFamily);
                w.WriteAttributeString("w", "hAnsi", W.NamespaceName, f.FontFamily);
                w.WriteAttributeString("w", "cs", W.NamespaceName, f.FontFamily);
                w.WriteEndElement();
            }
            if (f.Bold) WriteEmpty(w, "b");
            if (f.Italic) WriteEmpty(w, "i");
            if (f.Underline) WriteEmpty(w, "u");
            if (f.Strikeout) WriteEmpty(w, "strike");
            if (f.ForeColor.HasValue)
            {
                w.WriteStartElement("w", "color", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, ColorHex(f.ForeColor.Value));
                w.WriteEndElement();
            }
            if (f.BackColor.HasValue)
            {
                w.WriteStartElement("w", "shd", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, "clear");
                w.WriteAttributeString("w", "color", W.NamespaceName, "auto");
                w.WriteAttributeString("w", "fill", W.NamespaceName, ColorHex(f.BackColor.Value));
                w.WriteEndElement();
            }
            w.WriteStartElement("w", "sz", W.NamespaceName);
            w.WriteAttributeString("w", "val", W.NamespaceName, f.SizeHalfPoints.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteStartElement("w", "szCs", W.NamespaceName);
            w.WriteAttributeString("w", "val", W.NamespaceName, f.SizeHalfPoints.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            if (f.Script > 0)
            {
                w.WriteStartElement("w", "vertAlign", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, "superscript");
                w.WriteEndElement();
            }
            else if (f.Script < 0)
            {
                w.WriteStartElement("w", "vertAlign", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, "subscript");
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        private static void WriteEmpty(XmlWriter w, string name)
        {
            w.WriteStartElement("w", name, W.NamespaceName);
            w.WriteEndElement();
        }

        private static bool WriteImageRun(XmlWriter w, ImageInline img, List<ImagePart> images, ref int imageCounter)
        {
            ImagePart part = null;
            for (int i = 0; i < images.Count; i++)
            {
                if (images[i].Data == img.Data || (images[i].Data.Length == img.Data.Length && SameBytes(images[i].Data, img.Data)))
                {
                    part = images[i];
                    break;
                }
            }
            if (part == null) return false;

            imageCounter++;
            long cx = (long)Math.Max(1, img.WidthTwips) * 635L;
            long cy = (long)Math.Max(1, img.HeightTwips) * 635L;

            w.WriteStartElement("w", "r", W.NamespaceName);
            w.WriteStartElement("w", "drawing", W.NamespaceName);
            w.WriteStartElement("wp", "inline", WP.NamespaceName);
            w.WriteAttributeString("distT", "0");
            w.WriteAttributeString("distB", "0");
            w.WriteAttributeString("distL", "0");
            w.WriteAttributeString("distR", "0");
            w.WriteStartElement("wp", "extent", WP.NamespaceName);
            w.WriteAttributeString("cx", cx.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("cy", cy.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteStartElement("wp", "docPr", WP.NamespaceName);
            w.WriteAttributeString("id", imageCounter.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("name", "Imagen " + imageCounter);
            w.WriteEndElement();
            w.WriteStartElement("a", "graphic", A.NamespaceName);
            w.WriteStartElement("a", "graphicData", A.NamespaceName);
            w.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("pic", "pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("pic", "nvPicPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("pic", "cNvPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteAttributeString("id", imageCounter.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("name", "Imagen " + imageCounter);
            w.WriteEndElement();
            w.WriteStartElement("pic", "cNvPicPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("pic", "blipFill", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("a", "blip", A.NamespaceName);
            w.WriteAttributeString("r", "embed", R.NamespaceName, part.RelId);
            w.WriteEndElement();
            w.WriteStartElement("a", "stretch", A.NamespaceName);
            w.WriteStartElement("a", "fillRect", A.NamespaceName);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("pic", "spPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("a", "xfrm", A.NamespaceName);
            w.WriteStartElement("a", "off", A.NamespaceName);
            w.WriteAttributeString("x", "0");
            w.WriteAttributeString("y", "0");
            w.WriteEndElement();
            w.WriteStartElement("a", "ext", A.NamespaceName);
            w.WriteAttributeString("cx", cx.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("cy", cy.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("a", "prstGeom", A.NamespaceName);
            w.WriteAttributeString("prst", "rect");
            w.WriteStartElement("a", "avLst", A.NamespaceName);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            return true;
        }

        private static void WriteFloatingImageRun(XmlWriter w, FloatingObject floating,
            List<ImagePart> images, ref int imageCounter)
        {
            ImageInline img = floating.Image;
            ImagePart part = null;
            for (int i = 0; i < images.Count; i++)
            {
                if (images[i].Data == img.Data ||
                    (images[i].Data.Length == img.Data.Length && SameBytes(images[i].Data, img.Data)))
                {
                    part = images[i];
                    break;
                }
            }
            if (part == null) return;
            imageCounter++;
            long cx = (long)Math.Max(1, floating.WidthTwips) * 635L;
            long cy = (long)Math.Max(1, floating.HeightTwips) * 635L;
            long px = (long)floating.OffsetXTwips * 635L;
            long py = (long)floating.OffsetYTwips * 635L;
            w.WriteStartElement("w", "r", W.NamespaceName);
            w.WriteStartElement("w", "drawing", W.NamespaceName);
            w.WriteStartElement("wp", "anchor", WP.NamespaceName);
            w.WriteAttributeString("distT", "0");
            w.WriteAttributeString("distB", "0");
            w.WriteAttributeString("distL", "114300");
            w.WriteAttributeString("distR", "114300");
            w.WriteAttributeString("simplePos", "0");
            w.WriteAttributeString("relativeHeight", (imageCounter + 250).ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("behindDoc", floating.BehindText || floating.Wrap == TextWrapMode.Behind ? "1" : "0");
            w.WriteAttributeString("locked", "0");
            w.WriteAttributeString("layoutInCell", "1");
            w.WriteAttributeString("allowOverlap", "1");
            w.WriteStartElement("wp", "simplePos", WP.NamespaceName);
            w.WriteAttributeString("x", "0");
            w.WriteAttributeString("y", "0");
            w.WriteEndElement();
            w.WriteStartElement("wp", "positionH", WP.NamespaceName);
            w.WriteAttributeString("relativeFrom", "page");
            w.WriteStartElement("wp", "posOffset", WP.NamespaceName);
            w.WriteString(px.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("wp", "positionV", WP.NamespaceName);
            w.WriteAttributeString("relativeFrom", "page");
            w.WriteStartElement("wp", "posOffset", WP.NamespaceName);
            w.WriteString(py.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("wp", "extent", WP.NamespaceName);
            w.WriteAttributeString("cx", cx.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("cy", cy.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteStartElement("wp", "effectExtent", WP.NamespaceName);
            w.WriteAttributeString("l", "0");
            w.WriteAttributeString("t", "0");
            w.WriteAttributeString("r", "0");
            w.WriteAttributeString("b", "0");
            w.WriteEndElement();
            if (floating.Wrap == TextWrapMode.TopAndBottom)
            {
                w.WriteStartElement("wp", "wrapTopAndBottom", WP.NamespaceName);
                w.WriteEndElement();
            }
            else if (floating.Wrap == TextWrapMode.Square || floating.Wrap == TextWrapMode.Tight || floating.Wrap == TextWrapMode.Through)
            {
                w.WriteStartElement("wp", "wrapSquare", WP.NamespaceName);
                w.WriteAttributeString("wrapText", "bothSides");
                w.WriteEndElement();
            }
            else
            {
                w.WriteStartElement("wp", "wrapNone", WP.NamespaceName);
                w.WriteEndElement();
            }
            w.WriteStartElement("wp", "docPr", WP.NamespaceName);
            w.WriteAttributeString("id", imageCounter.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("name", "Imagen flotante " + imageCounter);
            w.WriteEndElement();
            w.WriteStartElement("a", "graphic", A.NamespaceName);
            w.WriteStartElement("a", "graphicData", A.NamespaceName);
            w.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("pic", "pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("pic", "nvPicPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("pic", "cNvPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteAttributeString("id", imageCounter.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("name", "Imagen flotante " + imageCounter);
            w.WriteEndElement();
            w.WriteStartElement("pic", "cNvPicPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("pic", "blipFill", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            w.WriteStartElement("a", "blip", A.NamespaceName);
            w.WriteAttributeString("r", "embed", R.NamespaceName, part.RelId);
            w.WriteEndElement();
            w.WriteStartElement("a", "stretch", A.NamespaceName);
            w.WriteStartElement("a", "fillRect", A.NamespaceName);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("pic", "spPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            if (floating.Rotation % 360 != 0)
                w.WriteAttributeString("rot", ((((long)floating.Rotation % 360 + 360) % 360) * 60000L).ToString(CultureInfo.InvariantCulture));
            w.WriteStartElement("a", "xfrm", A.NamespaceName);
            w.WriteStartElement("a", "off", A.NamespaceName);
            w.WriteAttributeString("x", "0");
            w.WriteAttributeString("y", "0");
            w.WriteEndElement();
            w.WriteStartElement("a", "ext", A.NamespaceName);
            w.WriteAttributeString("cx", cx.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("cy", cy.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("a", "prstGeom", A.NamespaceName);
            w.WriteAttributeString("prst", "rect");
            w.WriteStartElement("a", "avLst", A.NamespaceName);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteFloatingVmlRun(XmlWriter w, FloatingObject floating)
        {
            w.WriteStartElement("w", "r", W.NamespaceName);
            w.WriteStartElement("w", "pict", W.NamespaceName);
            w.WriteStartElement("v", "shape", VML.NamespaceName);
            w.WriteAttributeString("style", BuildVmlStyle(floating));
            w.WriteAttributeString("alt", floating.Kind == FloatingKind.Shape
                ? "MiniWordShape:" + floating.ShapeKind.ToString() : "MiniWordTextBox");
            w.WriteAttributeString("fillcolor", floating.Fill.HasValue ? "#" + ColorHex(floating.Fill.Value) : "#F2F6FC");
            w.WriteAttributeString("strokecolor", floating.BorderColor.HasValue ? "#" + ColorHex(floating.BorderColor.Value) : "#5A697D");
            w.WriteAttributeString("strokeweight", Math.Max(1, floating.BorderSize) + "pt");
            if (floating.Kind == FloatingKind.Shape)
            {
                // A v:shape without geometry is only a rectangle to Word and
                // LibreOffice.  Emit an explicit VML path so triangles,
                // diamonds and arrows survive an external round-trip.
                w.WriteAttributeString("coordsize", "21600,21600");
                w.WriteAttributeString("path", BuildVmlShapePath(floating.ShapeKind));
            }
            if (floating.Kind == FloatingKind.TextBox || !string.IsNullOrEmpty(floating.Text))
            {
                w.WriteStartElement("v", "textbox", VML.NamespaceName);
                w.WriteStartElement("w", "txbxContent", W.NamespaceName);
                string[] lines = Normalize(floating.Text).Split('\n');
                for (int i = 0; i < lines.Length; i++)
                {
                    w.WriteStartElement("w", "p", W.NamespaceName);
                    w.WriteStartElement("w", "r", W.NamespaceName);
                    WriteRunProperties(w, floating.Format ?? new RunFormat("Calibri", 20));
                    w.WriteStartElement("w", "t", W.NamespaceName);
                    w.WriteAttributeString("xml", "space", null, "preserve");
                    w.WriteString(lines[i]);
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                }
                w.WriteEndElement();
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteFloatingVmlImageRun(XmlWriter w, FloatingObject floating,
            List<ImagePart> images, ref int imageCounter, List<EmbeddedPart> embedded)
        {
            ImagePart part = null;
            for (int i = 0; i < images.Count; i++)
            {
                if (images[i].Data == floating.Image.Data ||
                    (images[i].Data.Length == floating.Image.Data.Length && SameBytes(images[i].Data, floating.Image.Data)))
                {
                    part = images[i];
                    break;
                }
            }
            if (part == null) return;
            imageCounter++;
            w.WriteStartElement("w", "r", W.NamespaceName);
            w.WriteStartElement("w", "object", W.NamespaceName);
            w.WriteStartElement("v", "shape", VML.NamespaceName);
            w.WriteAttributeString("style", BuildVmlStyle(floating));
            w.WriteStartElement("v", "imagedata", VML.NamespaceName);
            w.WriteAttributeString("r", "id", R.NamespaceName, part.RelId);
            w.WriteEndElement();
            EmbeddedPart embeddedPart = null;
            if (embedded != null && floating.EmbeddedData != null)
            {
                for (int i = 0; i < embedded.Count; i++)
                    if (embedded[i].Data.Length == floating.EmbeddedData.Length && SameBytes(embedded[i].Data, floating.EmbeddedData))
                    {
                        embeddedPart = embedded[i];
                        break;
                    }
            }
            if (embeddedPart != null)
            {
                w.WriteStartElement("o", "OLEObject", OLE.NamespaceName);
                w.WriteAttributeString("Type", "Embed");
                w.WriteAttributeString("ProgID", string.IsNullOrEmpty(floating.ProgId) ? "Package" : floating.ProgId);
                w.WriteAttributeString("ShapeID", string.IsNullOrEmpty(floating.ShapeId) ? "_x0000_i1025" : floating.ShapeId);
                w.WriteAttributeString("r", "id", R.NamespaceName, embeddedPart.RelId);
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static string BuildVmlShapePath(FloatingShapeKind kind)
        {
            switch (kind)
            {
                case FloatingShapeKind.Ellipse:
                    return "m10800,0q10800,0,21600,10800q21600,21600,10800,21600q0,21600,0,10800q0,0,10800,0x e";
                case FloatingShapeKind.Triangle:
                    return "m10800,0l21600,21600l0,21600x e";
                case FloatingShapeKind.Star:
                    return "m10800,0l13140,7740l21180,7740l14820,12540l17100,20520l10800,15840l4500,20520l6780,12540l420,7740l8460,7740x e";
                case FloatingShapeKind.Arrow:
                    return "m0,7200l14400,7200l14400,0l21600,10800l14400,21600l14400,14400l0,14400x e";
                case FloatingShapeKind.ArrowLeft:
                    return "m21600,7200l7200,7200l7200,0l0,10800l7200,21600l7200,14400l21600,14400x e";
                case FloatingShapeKind.Diamond:
                    return "m10800,0l21600,10800l10800,21600l0,10800x e";
                case FloatingShapeKind.Heart:
                    return "m10800,21600q0,14400,0,9000q0,1800,7200,1800q10800,1800,10800,5400q10800,1800,14400,1800q21600,1800,21600,9000q21600,14400,10800,21600x e";
                default:
                    return "m0,0l21600,0l21600,21600l0,21600x e";
            }
        }

        private static string BuildVmlStyle(FloatingObject floating)
        {
            float left = floating.OffsetXTwips / 20f;
            float top = floating.OffsetYTwips / 20f;
            float width = floating.WidthTwips / 20f;
            float height = floating.HeightTwips / 20f;
            string position = floating.Wrap == TextWrapMode.Behind ? "absolute" : "absolute";
            string wrap = floating.Wrap == TextWrapMode.TopAndBottom ? "topAndBottom"
                : (floating.Wrap == TextWrapMode.Square ? "square" : "none");
            return "position:" + position + ";margin-left:" + left.ToString(CultureInfo.InvariantCulture)
                + "pt;margin-top:" + top.ToString(CultureInfo.InvariantCulture)
                + "pt;width:" + width.ToString(CultureInfo.InvariantCulture)
                + "pt;height:" + height.ToString(CultureInfo.InvariantCulture)
                + "pt;z-index:" + (floating.BehindText ? -1 : 1)
                + (floating.Rotation % 360 != 0
                    ? ";rotation:" + (((floating.Rotation % 360) + 360) % 360).ToString(CultureInfo.InvariantCulture) + "deg"
                    : string.Empty)
                + ";mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap-style:" + wrap;
        }

        private static bool SameBytes(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
            return true;
        }

        private static byte[] BuildNotesXml(List<DocumentNote> notes, bool endnote)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    string rootName = endnote ? "endnotes" : "footnotes";
                    string itemName = endnote ? "endnote" : "footnote";
                    w.WriteStartElement("w", rootName, W.NamespaceName);
                    string[] special = new string[] { "separator", "continuationSeparator" };
                    for (int i = 0; i < 2; i++)
                    {
                        w.WriteStartElement("w", itemName, W.NamespaceName);
                        w.WriteAttributeString("w", "type", W.NamespaceName, special[i]);
                        w.WriteAttributeString("w", "id", W.NamespaceName, (i == 0 ? -1 : 0).ToString(CultureInfo.InvariantCulture));
                        w.WriteStartElement("w", "p", W.NamespaceName);
                        w.WriteStartElement("w", "r", W.NamespaceName);
                        w.WriteStartElement("w", i == 0 ? "separator" : "continuationSeparator", W.NamespaceName);
                        w.WriteEndElement();
                        w.WriteEndElement();
                        w.WriteEndElement();
                        w.WriteEndElement();
                    }
                    for (int i = 0; i < notes.Count; i++)
                    {
                        w.WriteStartElement("w", itemName, W.NamespaceName);
                        w.WriteAttributeString("w", "id", W.NamespaceName, notes[i].Id.ToString(CultureInfo.InvariantCulture));
                        string[] lines = Normalize(notes[i].Text).Split('\n');
                        for (int line = 0; line < lines.Length; line++)
                        {
                            w.WriteStartElement("w", "p", W.NamespaceName);
                            w.WriteStartElement("w", "r", W.NamespaceName);
                            w.WriteStartElement("w", "t", W.NamespaceName);
                            w.WriteAttributeString("xml", "space", null, "preserve");
                            w.WriteString(lines[line]);
                            w.WriteEndElement();
                            w.WriteEndElement();
                            w.WriteEndElement();
                        }
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildCommentsXml(DocumentModel doc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("w", "comments", W.NamespaceName);
                    for (int i = 0; i < doc.Comments.Count; i++)
                    {
                        DocumentComment comment = doc.Comments[i];
                        w.WriteStartElement("w", "comment", W.NamespaceName);
                        w.WriteAttributeString("w", "id", W.NamespaceName, comment.Id.ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("w", "author", W.NamespaceName, comment.Author ?? string.Empty);
                        w.WriteAttributeString("w", "initials", W.NamespaceName, comment.Initials ?? string.Empty);
                        w.WriteAttributeString("w", "date", W.NamespaceName, comment.Date ?? string.Empty);
                        WriteCommentParagraph(w, comment.Text);
                        for (int r = 0; r < comment.Replies.Count; r++)
                            WriteCommentParagraph(w, comment.Replies[r].Text);
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static void WriteCommentParagraph(XmlWriter w, string text)
        {
            string[] lines = Normalize(text).Split('\n');
            for (int line = 0; line < lines.Length; line++)
            {
                w.WriteStartElement("w", "p", W.NamespaceName);
                w.WriteStartElement("w", "r", W.NamespaceName);
                w.WriteStartElement("w", "t", W.NamespaceName);
                w.WriteAttributeString("xml", "space", null, "preserve");
                w.WriteString(lines[line]);
                w.WriteEndElement();
                w.WriteEndElement();
                w.WriteEndElement();
            }
        }
        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        }

        private static byte[] BuildStylesXml()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("w", "styles", W.NamespaceName);
                    w.WriteStartElement("w", "docDefaults", W.NamespaceName);
                    w.WriteStartElement("w", "rPrDefault", W.NamespaceName);
                    w.WriteStartElement("w", "rPr", W.NamespaceName);
                    w.WriteStartElement("w", "rFonts", W.NamespaceName);
                    w.WriteAttributeString("w", "ascii", W.NamespaceName, "Calibri");
                    w.WriteAttributeString("w", "hAnsi", W.NamespaceName, "Calibri");
                    w.WriteAttributeString("w", "cs", W.NamespaceName, "Calibri");
                    w.WriteEndElement();
                    w.WriteStartElement("w", "sz", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, "22");
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteStartElement("w", "pPrDefault", W.NamespaceName);
                    w.WriteStartElement("w", "pPr", W.NamespaceName);
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();

                    WriteStyle(w, "Normal", "paragraph", true, null, null);
                    WriteStyle(w, "DefaultParagraphFont", "character", true, null, null);
                    WriteStyle(w, "Hyperlink", "character", false, null, null);
                    WriteStyle(w, "Title", "paragraph", false, "Normal", "Title");
                    WriteStyle(w, "Subtitle", "paragraph", false, "Normal", "Subtitle");
                    WriteStyle(w, "Heading1", "paragraph", false, "Normal", "heading 1");
                    WriteStyle(w, "Heading2", "paragraph", false, "Normal", "heading 2");
                    WriteStyle(w, "Heading3", "paragraph", false, "Normal", "heading 3");
                    WriteStyle(w, "Quote", "paragraph", false, "Normal", "Quote");
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static void WriteStyle(XmlWriter w, string id, string type, bool isDefault, string basedOn, string name)
        {
            w.WriteStartElement("w", "style", W.NamespaceName);
            w.WriteAttributeString("w", "type", W.NamespaceName, type);
            if (isDefault) w.WriteAttributeString("w", "default", W.NamespaceName, "1");
            w.WriteAttributeString("w", "styleId", W.NamespaceName, id);
            w.WriteStartElement("w", "name", W.NamespaceName);
            w.WriteAttributeString("w", "val", W.NamespaceName, name ?? id);
            w.WriteEndElement();
            if (basedOn != null)
            {
                w.WriteStartElement("w", "basedOn", W.NamespaceName);
                w.WriteAttributeString("w", "val", W.NamespaceName, basedOn);
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        private static byte[] BuildNumberingXml()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("w", "numbering", W.NamespaceName);
                    w.WriteStartElement("w", "abstractNum", W.NamespaceName);
                    w.WriteAttributeString("w", "abstractNumId", W.NamespaceName, "0");
                    w.WriteStartElement("w", "multiLevelType", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, "hybridMultilevel");
                    w.WriteEndElement();
                    string[] bullets = new string[] { "\u2022", "o", "\u25AA" };
                    int[] lefts = new int[] { 720, 1440, 2160 };
                    for (int level = 0; level < 3; level++)
                    {
                        w.WriteStartElement("w", "lvl", W.NamespaceName);
                        w.WriteAttributeString("w", "ilvl", W.NamespaceName, level.ToString(CultureInfo.InvariantCulture));
                        w.WriteStartElement("w", "start", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, "1");
                        w.WriteEndElement();
                        w.WriteStartElement("w", "numFmt", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, "bullet");
                        w.WriteEndElement();
                        w.WriteStartElement("w", "lvlText", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, bullets[level]);
                        w.WriteEndElement();
                        w.WriteStartElement("w", "lvlJc", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, "left");
                        w.WriteEndElement();
                        w.WriteStartElement("w", "pPr", W.NamespaceName);
                        w.WriteStartElement("w", "ind", W.NamespaceName);
                        w.WriteAttributeString("w", "left", W.NamespaceName, lefts[level].ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("w", "hanging", W.NamespaceName, "360");
                        w.WriteEndElement();
                        w.WriteEndElement();
                        w.WriteStartElement("w", "rPr", W.NamespaceName);
                        w.WriteStartElement("w", "rFonts", W.NamespaceName);
                        w.WriteAttributeString("w", "ascii", W.NamespaceName, "Calibri");
                        w.WriteAttributeString("w", "hAnsi", W.NamespaceName, "Calibri");
                        w.WriteEndElement();
                        w.WriteEndElement();
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                    w.WriteStartElement("w", "abstractNum", W.NamespaceName);
                    w.WriteAttributeString("w", "abstractNumId", W.NamespaceName, "1");
                    w.WriteStartElement("w", "multiLevelType", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, "hybridMultilevel");
                    w.WriteEndElement();
                    int[] numberLefts = new int[] { 720, 1440, 2160 };
                    for (int level = 0; level < 3; level++)
                    {
                        w.WriteStartElement("w", "lvl", W.NamespaceName);
                        w.WriteAttributeString("w", "ilvl", W.NamespaceName, level.ToString(CultureInfo.InvariantCulture));
                        w.WriteStartElement("w", "start", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, "1");
                        w.WriteEndElement();
                        w.WriteStartElement("w", "numFmt", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, "decimal");
                        w.WriteEndElement();
                        w.WriteStartElement("w", "lvlText", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, "%" + (level + 1) + ".");
                        w.WriteEndElement();
                        w.WriteStartElement("w", "lvlJc", W.NamespaceName);
                        w.WriteAttributeString("w", "val", W.NamespaceName, "left");
                        w.WriteEndElement();
                        w.WriteStartElement("w", "pPr", W.NamespaceName);
                        w.WriteStartElement("w", "ind", W.NamespaceName);
                        w.WriteAttributeString("w", "left", W.NamespaceName, numberLefts[level].ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("w", "hanging", W.NamespaceName, "360");
                        w.WriteEndElement();
                        w.WriteEndElement();
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                    w.WriteStartElement("w", "num", W.NamespaceName);
                    w.WriteAttributeString("w", "numId", W.NamespaceName, "1");
                    w.WriteStartElement("w", "abstractNumId", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, "0");
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteStartElement("w", "num", W.NamespaceName);
                    w.WriteAttributeString("w", "numId", W.NamespaceName, "2");
                    w.WriteStartElement("w", "abstractNumId", W.NamespaceName);
                    w.WriteAttributeString("w", "val", W.NamespaceName, "1");
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildCoreProps(DocumentModel doc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("cp", "coreProperties", CP.NamespaceName);
                    w.WriteAttributeString("xmlns", "dc", null, DC.NamespaceName);
                    w.WriteAttributeString("xmlns", "dcterms", null, DCTERMS.NamespaceName);
                    w.WriteAttributeString("xmlns", "xsi", null, XSI.NamespaceName);
                    w.WriteStartElement("dc", "title", DC.NamespaceName);
                    w.WriteString(doc.Title ?? string.Empty);
                    w.WriteEndElement();
                    w.WriteStartElement("dc", "creator", DC.NamespaceName);
                    w.WriteString(string.IsNullOrEmpty(doc.Author) ? "Mini Word" : doc.Author);
                    w.WriteEndElement();
                    w.WriteStartElement("cp", "lastModifiedBy", CP.NamespaceName);
                    w.WriteString("Mini Word");
                    w.WriteEndElement();
                    string now = DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
                    w.WriteStartElement("dcterms", "created", DCTERMS.NamespaceName);
                    w.WriteAttributeString("xsi", "type", XSI.NamespaceName, "dcterms:W3CDTF");
                    w.WriteString(now);
                    w.WriteEndElement();
                    w.WriteStartElement("dcterms", "modified", DCTERMS.NamespaceName);
                    w.WriteAttributeString("xsi", "type", XSI.NamespaceName, "dcterms:W3CDTF");
                    w.WriteString(now);
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static byte[] BuildAppProps()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = CreateXmlWriter(ms))
                {
                    w.WriteStartDocument(true);
                    w.WriteStartElement("Properties", EP.NamespaceName);
                    w.WriteElementString("Application", EP.NamespaceName, "Mini Word");
                    w.WriteElementString("AppVersion", EP.NamespaceName, "1.0");
                    w.WriteEndElement();
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static string ColorHex(Color c)
        {
            return c.R.ToString("X2", CultureInfo.InvariantCulture)
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
            byte[] documentXml;
            if (!zip.TryGetValue("word/document.xml", out documentXml))
                throw new InvalidDataException("El archivo no contiene word/document.xml.");

            XDocument document = LoadXml(documentXml);
            XDocument styles = null;
            XDocument numbering = null;
            byte[] temp;
            if (zip.TryGetValue("word/styles.xml", out temp)) styles = LoadXml(temp);
            if (zip.TryGetValue("word/numbering.xml", out temp)) numbering = LoadXml(temp);

            Dictionary<string, string> rels = new Dictionary<string, string>(StringComparer.Ordinal);
            if (zip.TryGetValue("word/_rels/document.xml.rels", out temp))
            {
                XDocument relDoc = LoadXml(temp);
                foreach (XElement rel in relDoc.Root.Elements(REL + "Relationship"))
                {
                    string id = (string)rel.Attribute("Id");
                    string target = (string)rel.Attribute("Target");
                    if (id != null && target != null) rels[id] = target;
                }
            }

            Reader reader = new Reader(zip, rels, styles, numbering);
            DocumentModel model = reader.ReadDocument(document);
            DetectUnsupported(document, model);
            DetectSecurityRisks(zip, model);
            ReadCoreProperties(zip, model);
            PreserveUnknownParts(zip, model);
            return model;
        }

        private static void DetectSecurityRisks(Dictionary<string, byte[]> zip, DocumentModel model)
        {
            if (zip == null || model == null) return;
            bool vba = false;
            bool activeX = false;
            bool embedded = false;
            foreach (string key in zip.Keys)
            {
                string path = key.Replace('\\', '/');
                if (path.IndexOf("/vbaProject.bin", StringComparison.OrdinalIgnoreCase) >= 0
                    || path.IndexOf("/vbaData.xml", StringComparison.OrdinalIgnoreCase) >= 0
                    || path.IndexOf("/vbaProjectSignature.bin", StringComparison.OrdinalIgnoreCase) >= 0)
                    vba = true;
                if (path.StartsWith("word/activeX/", StringComparison.OrdinalIgnoreCase))
                    activeX = true;
                if (path.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase))
                    embedded = true;
            }
            if (vba)
                model.SecurityWarnings.Add("El documento contiene macros VBA. No se ejecutan y no se conservan al guardar.");
            if (activeX)
                model.SecurityWarnings.Add("El documento contiene controles ActiveX. Se descartan y no se ejecutan.");
            if (embedded)
                model.SecurityWarnings.Add("El documento contiene objetos OLE incrustados. Se conservan sin ejecutarse.");
        }

        private static void PreserveUnknownParts(Dictionary<string, byte[]> zip, DocumentModel model)
        {
            if (zip == null || model == null) return;
            long total = 0;
            foreach (KeyValuePair<string, byte[]> entry in zip)
            {
                string path = entry.Key.Replace('\\', '/').TrimStart('/');
                if (string.IsNullOrEmpty(path) || entry.Value == null) continue;
                bool keep = path.StartsWith("word/theme/", StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith("word/glossary/", StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith("word/charts/", StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith("word/diagrams/", StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(path, "word/fontTable.xml", StringComparison.OrdinalIgnoreCase);
                if (!keep) continue;
                if (path.EndsWith("/vbaProject.bin", StringComparison.OrdinalIgnoreCase)
                    || path.EndsWith("/vbaData.xml", StringComparison.OrdinalIgnoreCase)
                    || path.EndsWith("/vbaProjectSignature.bin", StringComparison.OrdinalIgnoreCase))
                    continue;
                byte[] preservedBytes = entry.Value;
                if (path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        XDocument relDoc = LoadXml(preservedBytes);
                        if (relDoc.Root != null)
                        {
                            List<XElement> external = new List<XElement>();
                            foreach (XElement rel in relDoc.Root.Elements(REL + "Relationship"))
                            {
                                string mode = (string)rel.Attribute("TargetMode");
                                if (string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase))
                                    external.Add(rel);
                            }
                            for (int i = 0; i < external.Count; i++) external[i].Remove();
                            if (external.Count > 0)
                            {
                                using (MemoryStream sanitized = new MemoryStream())
                                {
                                    relDoc.Save(sanitized);
                                    preservedBytes = sanitized.ToArray();
                                }
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                total += preservedBytes.Length;
                if (total > 32L * 1024 * 1024) break;
                model.PreservedParts[path] = preservedBytes;
            }
            PreserveContentTypes(zip, model);
            PreserveRelationships(zip, model);
        }

        private static void PreserveContentTypes(Dictionary<string, byte[]> zip, DocumentModel model)
        {
            byte[] data;
            if (!zip.TryGetValue("[Content_Types].xml", out data)) return;
            try
            {
                XDocument types = LoadXml(data);
                if (types.Root == null) return;
                foreach (XElement item in types.Root.Elements(CT + "Default"))
                {
                    string extension = (string)item.Attribute("Extension");
                    string contentType = (string)item.Attribute("ContentType");
                    if (string.IsNullOrEmpty(extension) || string.IsNullOrEmpty(contentType)) continue;
                    bool used = false;
                    foreach (string path in model.PreservedParts.Keys)
                    {
                        if (string.Equals(Path.GetExtension(path).TrimStart('.'), extension,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            used = true;
                            break;
                        }
                    }
                    if (used)
                    {
                        PreservedContentType entry = new PreservedContentType();
                        entry.Extension = extension;
                        entry.ContentType = contentType;
                        model.PreservedContentTypes.Add(entry);
                    }
                }
                foreach (XElement item in types.Root.Elements(CT + "Override"))
                {
                    string partName = (string)item.Attribute("PartName");
                    string contentType = (string)item.Attribute("ContentType");
                    if (string.IsNullOrEmpty(partName) || string.IsNullOrEmpty(contentType)) continue;
                    string normalized = partName.TrimStart('/').Replace('\\', '/');
                    if (model.PreservedParts.ContainsKey(normalized))
                    {
                        PreservedContentType entry = new PreservedContentType();
                        entry.PartName = normalized;
                        entry.ContentType = contentType;
                        model.PreservedContentTypes.Add(entry);
                    }
                }
            }
            catch { }
        }

        private static void PreserveRelationships(Dictionary<string, byte[]> zip, DocumentModel model)
        {
            byte[] data;
            if (!zip.TryGetValue("word/_rels/document.xml.rels", out data)) return;
            try
            {
                XDocument relDoc = LoadXml(data);
                if (relDoc.Root == null) return;
                foreach (XElement item in relDoc.Root.Elements(REL + "Relationship"))
                {
                    string target = (string)item.Attribute("Target");
                    string mode = (string)item.Attribute("TargetMode");
                    string type = (string)item.Attribute("Type");
                    if (string.IsNullOrEmpty(target) || string.IsNullOrEmpty(type)) continue;
                    string typeKey = type.ToLowerInvariant();
                    if (typeKey.IndexOf("vba", StringComparison.Ordinal) >= 0
                        || typeKey.IndexOf("activex", StringComparison.Ordinal) >= 0
                        || typeKey.IndexOf("control", StringComparison.Ordinal) >= 0)
                    {
                        string warning = "El documento contiene relaciones de macros/ActiveX. Se descartan.";
                        if (!model.SecurityWarnings.Contains(warning)) model.SecurityWarnings.Add(warning);
                        continue;
                    }
                    bool external = string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase);
                    if (external && typeKey.IndexOf("hyperlink", StringComparison.Ordinal) < 0)
                    {
                        string warning = "El documento contiene relaciones externas que no se descargan y se descartan.";
                        if (!model.SecurityWarnings.Contains(warning)) model.SecurityWarnings.Add(warning);
                        continue;
                    }
                    string normalized = target;
                    if (!external)
                    {
                        normalized = target.Replace('\\', '/');
                        if (normalized.StartsWith("../", StringComparison.Ordinal)) normalized = normalized.Substring(3);
                        if (!normalized.StartsWith("word/", StringComparison.OrdinalIgnoreCase))
                            normalized = "word/" + normalized.TrimStart('/');
                    }
                    if (!external && !model.PreservedParts.ContainsKey(normalized)) continue;
                    PreservedRelationship entry = new PreservedRelationship();
                    entry.Type = type;
                    entry.Target = target;
                    entry.TargetMode = mode ?? string.Empty;
                    model.PreservedRelationships.Add(entry);
                }
            }
            catch { }
        }

        private static void DetectUnsupported(XDocument document, DocumentModel model)
        {
            AddUnsupported(document, model, "sdt", "Controles de contenido");
            AddUnsupported(document, model, "ins", "Control de cambios (inserciones)");
            AddUnsupported(document, model, "del", "Control de cambios (eliminaciones)");
            AddUnsupported(document, model, "moveFrom", "Control de cambios (movimientos)");
            AddUnsupported(document, model, "moveTo", "Control de cambios (movimientos)");
            AddUnsupported(document, model, "fldSimple", "Campos e índices");
            AddUnsupported(document, model, "instrText", "Campos e índices");
            AddUnsupported(document, model, "fldChar", "Campos e índices");
            AddUnsupported(document, model, "oMath", "Ecuaciones");
            AddUnsupported(document, model, "oMathPara", "Ecuaciones");
            AddUnsupported(document, model, "chart", "Gráficos");
            AddUnsupported(document, model, "altChunk", "Contenido alternativo");
            AddUnsupported(document, model, "pgBorders", "Bordes de página");
            AddUnsupported(document, model, "background", "Fondo de página");
            AddUnsupported(document, model, "customXml", "XML personalizado");
        }

        private static void AddUnsupported(XDocument document, DocumentModel model,
            string localName, string description)
        {
            if (document == null || document.Root == null) return;
            foreach (XElement element in document.Descendants())
            {
                if (!string.Equals(element.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!model.UnsupportedFeatures.Contains(description))
                    model.UnsupportedFeatures.Add(description);
                return;
            }
        }

        private static void ReadCoreProperties(Dictionary<string, byte[]> zip, DocumentModel model)
        {
            byte[] data;
            if (!zip.TryGetValue("docProps/core.xml", out data)) return;
            try
            {
                XDocument core = LoadXml(data);
                XElement title = core.Root.Element(DC + "title");
                XElement creator = core.Root.Element(DC + "creator");
                if (title != null) model.Title = title.Value;
                if (creator != null) model.Author = creator.Value;
            }
            catch
            {
            }
        }

        private static XDocument LoadXml(byte[] data)
        {
            if (data == null || data.Length > 64 * 1024 * 1024)
                throw new InvalidDataException("XML del documento demasiado grande.");
            XDocument doc;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Prohibit;
            settings.XmlResolver = null;
            settings.MaxCharactersInDocument = 64 * 1024 * 1024;
            settings.MaxCharactersFromEntities = 0;
            using (MemoryStream ms = new MemoryStream(data, false))
            using (XmlReader reader = XmlReader.Create(ms, settings))
                doc = XDocument.Load(reader);
            return NormalizeNamespaces(doc);
        }

        private static XDocument NormalizeNamespaces(XDocument doc)
        {
            // ISO/IEC 29500 Strict uses different namespace URIs for the same
            // vocabulary.  Normalize them to the transitional names used by the
            // reader/writer, while preserving unrelated extension namespaces.
            Dictionary<string, XNamespace> map = new Dictionary<string, XNamespace>(StringComparer.Ordinal);
            map["http://purl.oclc.org/ooxml/wordprocessingml/main"] = W;
            map["http://purl.oclc.org/ooxml/officeDocument/relationships"] = R;
            map["http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"] = WP;
            map["http://purl.oclc.org/ooxml/drawingml/main"] = A;
            map["http://purl.oclc.org/ooxml/package/relationships"] = REL;
            map["http://purl.oclc.org/ooxml/officeDocument/math"] = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/math");
            bool changed = false;
            foreach (XElement element in doc.Descendants())
            {
                XNamespace ns;
                if (!map.TryGetValue(element.Name.NamespaceName, out ns)) continue;
                element.Name = ns + element.Name.LocalName;
                changed = true;
            }
            if (changed)
            {
                foreach (XElement element in doc.Descendants())
                {
                    List<XAttribute> attributes = new List<XAttribute>(element.Attributes());
                    for (int i = 0; i < attributes.Count; i++)
                    {
                        XNamespace ns;
                        if (map.TryGetValue(attributes[i].Name.NamespaceName, out ns))
                        {
                            string local = attributes[i].Name.LocalName;
                            string value = attributes[i].Value;
                            attributes[i].Remove();
                            element.SetAttributeValue(ns + local, value);
                        }
                    }
                }
            }
            return doc;
        }

        private sealed class Reader
        {
            private readonly Dictionary<string, byte[]> _zip;
            private readonly Dictionary<string, string> _rels;
            private readonly Dictionary<string, StyleDef> _styles = new Dictionary<string, StyleDef>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, RunFormat> _styleRunCache = new Dictionary<string, RunFormat>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, Paragraph> _styleParaCache = new Dictionary<string, Paragraph>(StringComparer.OrdinalIgnoreCase);
            private readonly RunFormat _defaultRun = new RunFormat();
            private readonly Paragraph _defaultPara = new Paragraph();
            private readonly Dictionary<int, Dictionary<int, NumLevel>> _numbering = new Dictionary<int, Dictionary<int, NumLevel>>();
            private readonly Dictionary<string, string> _defaultParagraphStyle = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<int, DocumentNote> _notesById = new Dictionary<int, DocumentNote>();
            private readonly Dictionary<int, DocumentComment> _commentsById = new Dictionary<int, DocumentComment>();

            public Reader(Dictionary<string, byte[]> zip, Dictionary<string, string> rels, XDocument styles, XDocument numbering)
            {
                _zip = zip;
                _rels = rels;
                if (styles != null) ParseStyles(styles);
                if (numbering != null) ParseNumbering(numbering);
            }

            private sealed class StyleDef
            {
                public string Id;
                public string Type;
                public string BasedOn;
                public XElement PPr;
                public XElement RPr;
            }

            private sealed class NumLevel
            {
                public string Format = "bullet";
                public int Left;
                public int Hanging;
            }

            private void ParseStyles(XDocument doc)
            {
                XElement root = doc.Root;
                if (root == null) return;
                XElement defaults = root.Element(W + "docDefaults");
                if (defaults != null)
                {
                    XElement rPr = defaults.Element(W + "rPrDefault");
                    if (rPr != null)
                    {
                        XElement props = rPr.Element(W + "rPr");
                        if (props != null) ApplyRunProperties(_defaultRun, props);
                    }
                    XElement pPr = defaults.Element(W + "pPrDefault");
                    if (pPr != null)
                    {
                        XElement props = pPr.Element(W + "pPr");
                        if (props != null) ApplyParagraphProperties(_defaultPara, props);
                    }
                }

                foreach (XElement style in root.Elements(W + "style"))
                {
                    StyleDef def = new StyleDef();
                    def.Id = (string)style.Attribute(W + "styleId");
                    def.Type = (string)style.Attribute(W + "type");
                    XElement based = style.Element(W + "basedOn");
                    if (based != null) def.BasedOn = (string)based.Attribute(W + "val");
                    def.PPr = style.Element(W + "pPr");
                    def.RPr = style.Element(W + "rPr");
                    if (def.Id != null) _styles[def.Id] = def;

                    string isDefault = (string)style.Attribute(W + "default");
                    if (isDefault == "1" || string.Equals(isDefault, "true", StringComparison.OrdinalIgnoreCase))
                    {
                        if (def.Type == "paragraph" && !_defaultParagraphStyle.ContainsKey("paragraph"))
                            _defaultParagraphStyle["paragraph"] = def.Id;
                        if (def.Type == "character" && !_defaultParagraphStyle.ContainsKey("character"))
                            _defaultParagraphStyle["character"] = def.Id;
                    }
                }
            }

            private void ParseNumbering(XDocument doc)
            {
                XElement root = doc.Root;
                if (root == null) return;
                Dictionary<int, Dictionary<int, NumLevel>> abstracts = new Dictionary<int, Dictionary<int, NumLevel>>();
                foreach (XElement abs in root.Elements(W + "abstractNum"))
                {
                    int absId;
                    if (!int.TryParse((string)abs.Attribute(W + "abstractNumId"), out absId)) continue;
                    Dictionary<int, NumLevel> levels = new Dictionary<int, NumLevel>();
                    foreach (XElement lvl in abs.Elements(W + "lvl"))
                    {
                        int ilvl;
                        if (!int.TryParse((string)lvl.Attribute(W + "ilvl"), out ilvl)) continue;
                        NumLevel nl = new NumLevel();
                        XElement fmt = lvl.Element(W + "numFmt");
                        if (fmt != null) nl.Format = (string)fmt.Attribute(W + "val") ?? "bullet";
                        XElement pPr = lvl.Element(W + "pPr");
                        if (pPr != null)
                        {
                            XElement ind = pPr.Element(W + "ind");
                            if (ind != null)
                            {
                                nl.Left = ReadIntAttr(ind, "left", "start");
                                nl.Hanging = ReadIntAttr(ind, "hanging");
                            }
                        }
                        levels[ilvl] = nl;
                    }
                    abstracts[absId] = levels;
                }

                foreach (XElement num in root.Elements(W + "num"))
                {
                    int numId;
                    if (!int.TryParse((string)num.Attribute(W + "numId"), out numId)) continue;
                    XElement abs = num.Element(W + "abstractNumId");
                    int absId;
                    if (abs != null && int.TryParse((string)abs.Attribute(W + "val"), out absId) && abstracts.ContainsKey(absId))
                        _numbering[numId] = abstracts[absId];
                }
            }

            private void ReadNotes(DocumentModel model)
            {
                ReadNoteFile("word/footnotes.xml", "footnotes", false, model);
                ReadNoteFile("word/endnotes.xml", "endnotes", true, model);
                ReadCommentFile(model);
            }

            private void ReadNoteFile(string path, string rootName, bool endnote, DocumentModel model)
            {
                byte[] data;
                if (!_zip.TryGetValue(path, out data)) return;
                try
                {
                    XDocument doc = LoadXml(data);
                    XElement root = doc.Root;
                    if (root == null) return;
                    foreach (XElement note in root.Elements())
                    {
                        if (note.Name.LocalName != (endnote ? "endnote" : "footnote")) continue;
                        int id;
                        if (!int.TryParse((string)note.Attribute(W + "id"), out id) || id <= 0) continue;
                        if (string.Equals((string)note.Attribute(W + "type"), "separator", StringComparison.OrdinalIgnoreCase)
                            || string.Equals((string)note.Attribute(W + "type"), "continuationSeparator", StringComparison.OrdinalIgnoreCase)) continue;
                        DocumentNote item = new DocumentNote();
                        item.Id = id;
                        item.Endnote = endnote;
                        StringBuilder text = new StringBuilder();
                        foreach (XElement p in note.Descendants(W + "p"))
                        {
                            if (text.Length > 0) text.Append('\n');
                            text.Append(CollectText(p));
                        }
                        item.Text = text.ToString();
                        if (endnote) model.Endnotes.Add(item); else model.Footnotes.Add(item);
                        _notesById[id] = item;
                    }
                }
                catch { }
            }

            private void ReadCommentFile(DocumentModel model)
            {
                byte[] data;
                if (!_zip.TryGetValue("word/comments.xml", out data)) return;
                try
                {
                    XDocument doc = LoadXml(data);
                    XElement root = doc.Root;
                    if (root == null) return;
                    foreach (XElement comment in root.Elements(W + "comment"))
                    {
                        int id;
                        if (!int.TryParse((string)comment.Attribute(W + "id"), out id)) continue;
                        DocumentComment item = new DocumentComment();
                        item.Id = id;
                        item.Author = (string)comment.Attribute(W + "author") ?? string.Empty;
                        item.Initials = (string)comment.Attribute(W + "initials") ?? string.Empty;
                        item.Date = (string)comment.Attribute(W + "date") ?? string.Empty;
                        bool firstParagraph = true;
                        foreach (XElement p in comment.Descendants(W + "p"))
                        {
                            string line = CollectText(p);
                            if (firstParagraph)
                            {
                                item.Text = line;
                                firstParagraph = false;
                            }
                            else
                            {
                                item.Replies.Add(new CommentReply
                                {
                                    Text = line,
                                    Author = item.Author,
                                    Date = item.Date
                                });
                            }
                        }                        model.Comments.Add(item);
                        _commentsById[id] = item;
                    }
                }
                catch { }
            }

            public DocumentModel ReadDocument(XDocument document)
            {
                DocumentModel model = new DocumentModel();
                XElement body = document.Root.Element(W + "body");
                if (body == null) return model;
                ReadNotes(model);

                List<SectionMark> marks = new List<SectionMark>();
                foreach (XElement child in body.Elements())
                    ReadBlock(child, model, marks, 0);

                if (model.Paragraphs.Count == 0)
                    model.Paragraphs.Add(new Paragraph());

                XElement lastSect = body.Element(W + "sectPr");
                BuildSections(model, marks, lastSect);
                return model;
            }

            private sealed class SectionMark
            {
                public int ParagraphIndex;
                public XElement Properties;
            }

            private const int MaxXmlDepth = 128;

            private void ReadBlock(XElement child, DocumentModel model, List<SectionMark> marks, int depth)
            {
                if (depth > MaxXmlDepth)
                    throw new InvalidDataException("El DOCX contiene un anidamiento XML excesivo.");
                if (child.Name == W + "p")
                {
                    model.Paragraphs.Add(ReadParagraph(child));
                    XElement pPr = child.Element(W + "pPr");
                    XElement sectPr = pPr != null ? pPr.Element(W + "sectPr") : null;
                    if (sectPr != null)
                    {
                        SectionMark mark = new SectionMark();
                        mark.ParagraphIndex = model.Paragraphs.Count - 1;
                        mark.Properties = sectPr;
                        marks.Add(mark);
                    }
                }
                else if (child.Name == W + "tbl")
                {
                    ReadTable(child, model);
                }
                else if (child.Name == W + "sectPr")
                {
                    // The body-level section is applied after all content is read.
                }
                else if (child.Name == W + "sdt" || child.Name == W + "customXml" || child.Name == W + "smartTag")
                {
                    XElement content = child.Element(W + "sdtContent");
                    if (content == null) content = child;
                    foreach (XElement inner in content.Elements()) ReadBlock(inner, model, marks, depth + 1);
                }
            }

            private void BuildSections(DocumentModel model, List<SectionMark> marks, XElement bodySectPr)
            {
                model.Sections.Clear();
                int start = 0;
                for (int i = 0; i < marks.Count; i++)
                {
                    DocumentSection section = DocumentSection.FromDocument(model, start);
                    ReadSectionInto(marks[i].Properties, section);
                    if (model.Sections.Count > 0)
                    {
                        DocumentSection previous = model.Sections[model.Sections.Count - 1];
                        if (string.IsNullOrEmpty(section.HeaderText)) section.HeaderText = previous.HeaderText;
                        if (string.IsNullOrEmpty(section.FooterText)) section.FooterText = previous.FooterText;
                    }
                    model.Sections.Add(section);
                    start = marks[i].ParagraphIndex + 1;
                }
                if (bodySectPr != null || model.Sections.Count == 0 || start < model.Paragraphs.Count)
                {
                    DocumentSection section = DocumentSection.FromDocument(model, start);
                    if (bodySectPr != null) ReadSectionInto(bodySectPr, section);
                    if (model.Sections.Count > 0)
                    {
                        DocumentSection previous = model.Sections[model.Sections.Count - 1];
                        if (string.IsNullOrEmpty(section.HeaderText)) section.HeaderText = previous.HeaderText;
                        if (string.IsNullOrEmpty(section.FooterText)) section.FooterText = previous.FooterText;
                    }
                    model.Sections.Add(section);
                }
                if (model.Sections.Count > 0)
                {
                    int sectionIndex = 0;
                    for (int pi = 0; pi < model.Paragraphs.Count; pi++)
                    {
                        while (sectionIndex + 1 < model.Sections.Count &&
                            model.Sections[sectionIndex + 1].StartParagraph <= pi) sectionIndex++;
                        model.Paragraphs[pi].SectionIndex = sectionIndex;
                    }
                    DocumentSection first = model.Sections[0];
                    model.PageWidthTwips = first.PageWidthTwips;
                    model.PageHeightTwips = first.PageHeightTwips;
                    model.MarginLeftTwips = first.MarginLeftTwips;
                    model.MarginRightTwips = first.MarginRightTwips;
                    model.MarginTopTwips = first.MarginTopTwips;
                    model.MarginBottomTwips = first.MarginBottomTwips;
                    model.HeaderText = first.HeaderText;
                    model.FooterText = first.FooterText;
                }
            }

            private Paragraph ReadParagraph(XElement element)
            {
                Paragraph p = new Paragraph();
                CopyParagraphFormat(_defaultPara, p);
                RunFormat baseRun = _defaultRun.Clone();

                XElement pPr = element.Element(W + "pPr");
                string pStyle = null;
                if (pPr != null)
                {
                    XElement style = pPr.Element(W + "pStyle");
                    if (style != null) pStyle = (string)style.Attribute(W + "val");
                }
                if (pStyle == null && _defaultParagraphStyle.ContainsKey("paragraph"))
                    pStyle = _defaultParagraphStyle["paragraph"];

                if (!string.IsNullOrEmpty(pStyle))
                {
                    p.StyleName = DocumentStyles.NameFromId(pStyle);
                    Paragraph stylePara = GetParagraphStyle(pStyle);
                    if (stylePara != null) CopyParagraphFormat(stylePara, p);
                    RunFormat styleRun = GetRunStyle(pStyle);
                    if (styleRun != null) baseRun = styleRun.Clone();
                }

                if (pPr != null) ApplyParagraphProperties(p, pPr);
                if (element.Descendants(W + "br").Any(delegate(XElement br) {
                    return string.Equals((string)br.Attribute(W + "type"), "page", StringComparison.OrdinalIgnoreCase);
                })) p.PageBreakBefore = true;

                XElement numPr = pPr != null ? pPr.Element(W + "numPr") : null;
                bool hasNumbering = false;
                if (numPr != null)
                {
                    int numId = ReadIntElement(numPr, "numId");
                    int ilvl = ReadIntElement(numPr, "ilvl");
                    hasNumbering = ApplyNumbering(p, numId, ilvl, baseRun);
                }

                ReadInlines(element, p, baseRun, pStyle);

                if (p.Inlines.Count == 0)
                    p.Inlines.Add(new TextInline(string.Empty, baseRun.Clone()));

                bool preserve = false;
                foreach (XElement node in element.Descendants())
                {
                    string name = node.Name.LocalName;
                    if (string.Equals(name, "fldSimple", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(name, "instrText", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(name, "fldChar", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(name, "sdt", StringComparison.OrdinalIgnoreCase))
                    {
                        preserve = true;
                        break;
                    }
                }
                if (p.HasModeledField) preserve = false;
                if (preserve)
                {
                    p.PreservedXml = element.ToString(SaveOptions.DisableFormatting);
                    StringBuilder modelText = new StringBuilder();
                    for (int i = 0; i < p.Inlines.Count; i++)
                    {
                        TextInline text = p.Inlines[i] as TextInline;
                        if (text != null) modelText.Append(text.Text);
                    }
                    p.PreservedText = modelText.ToString();
                    p.PreservedFormatKey = ParagraphFormatKey(p);
                }

                return p;
            }

            private bool ApplyNumbering(Paragraph p, int numId, int ilvl, RunFormat baseRun)
            {
                Dictionary<int, NumLevel> levels;
                if (!_numbering.TryGetValue(numId, out levels)) return false;
                NumLevel level;
                if (!levels.TryGetValue(ilvl, out level)) level = null;
                string format = level != null ? level.Format : "bullet";
                bool bullet = string.Equals(format, "bullet", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(format);

                if (level != null)
                {
                    if (level.Left > 0 && p.LeftIndentTwips == 0) p.LeftIndentTwips = level.Left;
                    if (level.Hanging > 0 && p.FirstLineIndentTwips == 0) p.FirstLineIndentTwips = -level.Hanging;
                }

                if (bullet)
                {
                    p.Bullet = true;
                    p.BulletLevel = Math.Min(8, Math.Max(0, ilvl));
                    return true;
                }

                p.Numbered = true;
                p.NumberLevel = Math.Min(8, Math.Max(0, ilvl));
                return true;
            }

            private void ReadInlines(XElement parent, Paragraph p, RunFormat baseRun, string paragraphStyle)
            {
                ReadInlines(parent, p, baseRun, paragraphStyle, RevisionKind.None, string.Empty, string.Empty, 0);
            }

            private void ReadInlines(XElement parent, Paragraph p, RunFormat baseRun, string paragraphStyle,
                RevisionKind revision, string author, string date, int depth)
            {
                if (depth > MaxXmlDepth)
                    throw new InvalidDataException("El DOCX contiene un anidamiento XML excesivo.");
                foreach (XElement child in parent.Elements())
                {
                    if (child.Name == W + "r")
                    {
                        ReadRun(child, p, baseRun, paragraphStyle, revision, author, date);
                    }
                    else if (child.Name == W + "hyperlink")
                    {
                        RunFormat link = baseRun.Clone();
                        link.ForeColor = Color.FromArgb(0, 0, 238);
                        link.Underline = true;
                        string rid = (string)child.Attribute(R + "id");
                        if (rid != null && _rels.ContainsKey(rid)) link.LinkUrl = _rels[rid];
                        foreach (XElement run in child.Elements(W + "r"))
                            ReadRun(run, p, link, paragraphStyle, revision, author, date);
                    }
                    else if (child.Name == W + "ins" || child.Name == W + "moveTo")
                    {
                        ReadInlines(child, p, baseRun, paragraphStyle, RevisionKind.Insert,
                            (string)child.Attribute(W + "author") ?? "Autor",
                            (string)child.Attribute(W + "date") ?? string.Empty, depth + 1);
                    }
                    else if (child.Name == W + "del" || child.Name == W + "moveFrom")
                    {
                        ReadInlines(child, p, baseRun, paragraphStyle, RevisionKind.Delete,
                            (string)child.Attribute(W + "author") ?? "Autor",
                            (string)child.Attribute(W + "date") ?? string.Empty, depth + 1);
                    }
                    else if (child.Name == W + "fldSimple")
                    {
                        string instr = (string)child.Attribute(W + "instr") ?? string.Empty;
                        FieldKind field = FieldText.ParseInstruction(instr);
                        if (field != FieldKind.None)
                        {
                            TextInline run = NewRun(CollectSimpleFieldText(child), baseRun.Clone(), revision, author, date);
                            run.Field = field;
                            run.FieldInstruction = instr;
                            p.Inlines.Add(run);
                            p.HasModeledField = true;
                        }
                        else
                        {
                            ReadInlines(child, p, baseRun, paragraphStyle, revision, author, date, depth + 1);
                        }
                    }
                    else if (child.Name == W + "smartTag" || child.Name == W + "sdt"
                        || child.Name == W + "sdtContent")
                    {
                        ReadInlines(child, p, baseRun, paragraphStyle, revision, author, date, depth + 1);
                    }
                    else if (child.Name == W + "bookmarkStart" || child.Name == W + "bookmarkEnd"
                        || child.Name == W + "proofErr" || child.Name == W + "commentRangeStart"
                        || child.Name == W + "commentRangeEnd")
                    {
                        // ignored
                    }
                }
            }

            private static string CollectSimpleFieldText(XElement element)
            {
                StringBuilder text = new StringBuilder();
                foreach (XElement node in element.Descendants())
                {
                    if (node.Name == W + "t" || node.Name == W + "delText") text.Append(node.Value);
                    else if (node.Name == W + "tab") text.Append('\t');
                    else if (node.Name == W + "br" || node.Name == W + "cr") text.Append('\n');
                }
                return text.ToString();
            }

            private static TextInline NewRun(string text, RunFormat format, RevisionKind revision,
                string author, string date)
            {
                TextInline run = new TextInline(text, format);
                run.Revision = revision;
                run.RevisionAuthor = author ?? string.Empty;
                run.RevisionDate = date ?? string.Empty;
                return run;
            }
            private void ReadRun(XElement run, Paragraph p, RunFormat baseRun, string paragraphStyle)
            {
                ReadRun(run, p, baseRun, paragraphStyle, RevisionKind.None, string.Empty, string.Empty);
            }

            private void ReadRun(XElement run, Paragraph p, RunFormat baseRun, string paragraphStyle,
                RevisionKind revision, string author, string date)
            {
                RunFormat fmt = baseRun.Clone();
                XElement rPr = run.Element(W + "rPr");
                string rStyle = null;
                if (rPr != null)
                {
                    XElement rs = rPr.Element(W + "rStyle");
                    if (rs != null) rStyle = (string)rs.Attribute(W + "val");
                    if (!string.IsNullOrEmpty(rStyle))
                    {
                        ApplyCharacterStyle(fmt, rStyle, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
                    }
                    ApplyRunProperties(fmt, rPr);
                }

                StringBuilder text = new StringBuilder();
                foreach (XElement child in run.Elements())
                {
                    if (child.Name == W + "t" || child.Name == W + "delText")
                    {
                        text.Append(child.Value);
                    }
                    else if (child.Name == W + "tab")
                    {
                        text.Append('\t');
                    }
                    else if (child.Name == W + "br" || child.Name == W + "cr")
                    {
                        text.Append('\n');
                    }
                    else if (child.Name == W + "noBreakHyphen")
                    {
                        text.Append('-');
                    }
                    else if (child.Name == W + "sym")
                    {
                        string ch = (string)child.Attribute(W + "char");
                        if (!string.IsNullOrEmpty(ch))
                        {
                            int code;
                            if (int.TryParse(ch, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                                text.Append((char)code);
                        }
                    }
                    else if (child.Name == W + "footnoteReference" || child.Name == W + "endnoteReference")
                    {
                        int id;
                        if (int.TryParse((string)child.Attribute(W + "id"), out id) && _notesById.ContainsKey(id))
                        {
                            if (text.Length > 0)
                            {
                                p.Inlines.Add(NewRun(text.ToString(), fmt.Clone(), revision, author, date));
                                text.Length = 0;
                            }
                            RunFormat refFormat = fmt.Clone();
                            refFormat.Script = 1;
                            p.Inlines.Add(new NoteInline(_notesById[id], refFormat));
                        }
                    }
                    else if (child.Name == W + "commentReference")
                    {
                        int id;
                        if (int.TryParse((string)child.Attribute(W + "id"), out id) && _commentsById.ContainsKey(id))
                        {
                            if (text.Length > 0)
                            {
                                p.Inlines.Add(NewRun(text.ToString(), fmt.Clone(), revision, author, date));
                                text.Length = 0;
                            }
                            p.Inlines.Add(new CommentInline(_commentsById[id], fmt.Clone()));
                        }
                    }
                    else if (child.Name == W + "drawing")
                    {
                        if (TryReadFloatingDrawing(child, p, fmt)) continue;
                        ImageInline img = ReadDrawing(child);
                        if (img != null)
                        {
                            if (text.Length > 0)
                            {
                                p.Inlines.Add(NewRun(text.ToString(), fmt.Clone(), revision, author, date));
                                text.Length = 0;
                            }
                            img.Format = fmt.Clone();
                            p.Inlines.Add(img);
                        }
                    }
                    else if (child.Name == W + "pict")
                    {
                        if (TryReadVmlFloating(child, p, fmt)) continue;
                        ImageInline img = ReadVmlImage(child);
                        if (img != null)
                        {
                            if (text.Length > 0)
                            {
                                p.Inlines.Add(NewRun(text.ToString(), fmt.Clone(), revision, author, date));
                                text.Length = 0;
                            }
                            img.Format = fmt.Clone();
                            p.Inlines.Add(img);
                        }
                    }
                    else if (child.Name.LocalName == "AlternateContent")
                    {
                        bool handled = false;
                        foreach (XElement drawing in child.Descendants(W + "drawing"))
                            if (TryReadFloatingDrawing(drawing, p, fmt)) handled = true;
                        if (!handled)
                        {
                            foreach (XElement pict in child.Descendants(W + "pict"))
                                if (TryReadVmlFloating(pict, p, fmt)) handled = true;
                        }
                        if (handled)
                        {
                            if (text.Length > 0)
                            {
                                p.Inlines.Add(NewRun(text.ToString(), fmt.Clone(), revision, author, date));
                                text.Length = 0;
                            }
                        }
                    }
                    else if (child.Name == W + "object")
                    {
                        FloatingObject embedded = ReadEmbeddedObject(child, p, fmt);
                        if (embedded != null)
                        {
                            if (text.Length > 0)
                            {
                                p.Inlines.Add(NewRun(text.ToString(), fmt.Clone(), revision, author, date));
                                text.Length = 0;
                            }
                            p.FloatingObjects.Add(embedded);
                        }
                    }
                }

                if (text.Length > 0)
                    p.Inlines.Add(NewRun(text.ToString(), fmt.Clone(), revision, author, date));
            }

            private bool TryReadFloatingDrawing(XElement drawing, Paragraph p, RunFormat fmt)
            {
                XElement anchor = drawing.Element(WP + "anchor");
                if (anchor == null) return false;
                FloatingObject obj = new FloatingObject();
                obj.Format = fmt.Clone();
                XElement extent = anchor.Element(WP + "extent");
                ReadExtent(extent, obj);
                obj.OffsetXTwips = ReadPosition(anchor, "positionH");
                obj.OffsetYTwips = ReadPosition(anchor, "positionV");
                XElement rotationTransform = drawing.Descendants(A + "xfrm").FirstOrDefault();
                if (rotationTransform != null)
                {
                    long rotationValue;
                    if (long.TryParse((string)rotationTransform.Attribute("rot"), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out rotationValue))
                        obj.Rotation = (int)((rotationValue / 60000L) % 360L);
                }
                obj.BehindText = string.Equals((string)anchor.Attribute("behindDoc"), "1", StringComparison.OrdinalIgnoreCase)
                    || string.Equals((string)anchor.Attribute("behindDoc"), "true", StringComparison.OrdinalIgnoreCase);
                obj.Wrap = ReadWrap(anchor);

                XElement textBox = drawing.Descendants(WPS + "txbx").FirstOrDefault();
                if (textBox != null)
                {
                    XElement content = textBox.Element(W + "txbxContent");
                    obj.Kind = FloatingKind.TextBox;
                    obj.Text = content != null ? CollectBlockText(content) : string.Empty;
                    if (content != null)
                    {
                        XElement run = content.Descendants(W + "r").FirstOrDefault();
                        XElement runProps = run != null ? run.Element(W + "rPr") : null;
                        if (runProps != null) ApplyRunProperties(obj.Format, runProps);
                    }
                    ReadShapeFill(drawing, obj);
                    p.FloatingObjects.Add(obj);
                    return true;
                }

                XElement blip = drawing.Descendants(A + "blip").FirstOrDefault();
                if (blip != null)
                {
                    ImageInline image = ReadDrawing(drawing);
                    if (image == null) return false;
                    obj.Kind = FloatingKind.Image;
                    obj.Image = image;
                    obj.WidthTwips = image.WidthTwips;
                    obj.HeightTwips = image.HeightTwips;
                    p.FloatingObjects.Add(obj);
                    return true;
                }

                XElement shape = drawing.Descendants(WPS + "wsp").FirstOrDefault();
                if (shape != null)
                {
                    obj.Kind = FloatingKind.Shape;
                    XElement preset = shape.Descendants(A + "prstGeom").FirstOrDefault();
                    string presetName = preset != null ? (string)preset.Attribute("prst") : null;
                    if (string.Equals(presetName, "ellipse", StringComparison.OrdinalIgnoreCase)) obj.ShapeKind = FloatingShapeKind.Ellipse;
                    else if (string.Equals(presetName, "triangle", StringComparison.OrdinalIgnoreCase)) obj.ShapeKind = FloatingShapeKind.Triangle;
                    else if (string.Equals(presetName, "star5", StringComparison.OrdinalIgnoreCase)) obj.ShapeKind = FloatingShapeKind.Star;
                    else if (string.Equals(presetName, "rightArrow", StringComparison.OrdinalIgnoreCase)) obj.ShapeKind = FloatingShapeKind.Arrow;
                    else if (string.Equals(presetName, "leftArrow", StringComparison.OrdinalIgnoreCase)) obj.ShapeKind = FloatingShapeKind.ArrowLeft;
                    else if (string.Equals(presetName, "diamond", StringComparison.OrdinalIgnoreCase)) obj.ShapeKind = FloatingShapeKind.Diamond;
                    else if (string.Equals(presetName, "heart", StringComparison.OrdinalIgnoreCase)) obj.ShapeKind = FloatingShapeKind.Heart;
                    ReadShapeFill(drawing, obj);
                    XElement text = shape.Descendants(W + "t").FirstOrDefault();
                    obj.Text = text != null ? text.Value : string.Empty;
                    p.FloatingObjects.Add(obj);
                    return true;
                }
                return false;
            }

            private bool TryReadVmlFloating(XElement pict, Paragraph p, RunFormat fmt)
            {
                XElement shape = pict.Descendants(VML + "shape").FirstOrDefault();
                if (shape == null) return false;
                string alt = (string)shape.Attribute("alt") ?? string.Empty;
                XElement textBox = shape.Descendants(VML + "textbox").FirstOrDefault();
                // A shape without text is now written without a v:textbox
                // child (otherwise Word/LibreOffice draws an extra rectangle).
                if (textBox == null && !alt.StartsWith("MiniWordShape", StringComparison.Ordinal))
                    return false;
                FloatingObject obj = new FloatingObject();
                obj.Kind = alt.StartsWith("MiniWordShape", StringComparison.Ordinal)
                    ? FloatingKind.Shape : FloatingKind.TextBox;
                if (obj.Kind == FloatingKind.Shape)
                {
                    int colon = alt.IndexOf(':');
                    if (colon >= 0)
                    {
                        try { obj.ShapeKind = (FloatingShapeKind)Enum.Parse(typeof(FloatingShapeKind), alt.Substring(colon + 1), true); }
                        catch { obj.ShapeKind = FloatingShapeKind.Rectangle; }
                    }
                }
                obj.Format = fmt.Clone();
                XElement boxRun = null;
                if (textBox != null)
                {
                    foreach (XElement run in textBox.Descendants(W + "r"))
                    {
                        boxRun = run;
                        break;
                    }
                }
                if (boxRun != null)
                {
                    XElement boxProps = boxRun.Element(W + "rPr");
                    if (boxProps != null) ApplyRunProperties(obj.Format, boxProps);
                }
                obj.Text = textBox != null ? CollectBlockText(textBox) : string.Empty;
                string style = (string)shape.Attribute("style");
                int width = 1440, height = 720;
                if (!string.IsNullOrEmpty(style)) ParseCssSize(style, ref width, ref height);
                obj.WidthTwips = width;
                obj.HeightTwips = height;
                if (!string.IsNullOrEmpty(style))
                {
                    int offset;
                    if (TryCssSignedLength(style, "margin-left", out offset)) obj.OffsetXTwips = offset;
                    if (TryCssSignedLength(style, "margin-top", out offset)) obj.OffsetYTwips = offset;
                    int rotationAt = style.IndexOf("rotation:", StringComparison.OrdinalIgnoreCase);
                    if (rotationAt >= 0)
                    {
                        int valueAt = rotationAt + "rotation:".Length;
                        int valueEnd = valueAt;
                        while (valueEnd < style.Length && (char.IsDigit(style[valueEnd]) || style[valueEnd] == '-'))
                            valueEnd++;
                        int rotationValue;
                        if (valueEnd > valueAt && int.TryParse(style.Substring(valueAt, valueEnd - valueAt),
                            NumberStyles.Integer, CultureInfo.InvariantCulture, out rotationValue))
                            obj.Rotation = ((rotationValue % 360) + 360) % 360;
                    }
                    if (style.IndexOf("mso-wrap-style:topAndBottom", StringComparison.OrdinalIgnoreCase) >= 0)
                        obj.Wrap = TextWrapMode.TopAndBottom;
                    else if (style.IndexOf("mso-wrap-style:none", StringComparison.OrdinalIgnoreCase) >= 0)
                        obj.Wrap = style.IndexOf("z-index:-1", StringComparison.OrdinalIgnoreCase) >= 0
                            ? TextWrapMode.Behind : TextWrapMode.InFront;
                    else
                        obj.Wrap = TextWrapMode.Square;
                }
                string fillColor = (string)shape.Attribute("fillcolor");
                Color parsedColor;
                if (!string.IsNullOrEmpty(fillColor) && TryHexColor(fillColor, out parsedColor))
                    obj.Fill = parsedColor;
                string strokeColor = (string)shape.Attribute("strokecolor");
                if (!string.IsNullOrEmpty(strokeColor) && TryHexColor(strokeColor, out parsedColor))
                    obj.BorderColor = parsedColor;
                string strokeWeight = (string)shape.Attribute("strokeweight");
                if (!string.IsNullOrEmpty(strokeWeight))
                {
                    float weight;
                    if (float.TryParse(strokeWeight.TrimEnd('p', 't', 'P', 'T'), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out weight))
                        obj.BorderSize = Math.Max(1, (int)Math.Round(weight));
                }
                obj.BehindText = style != null
                    && style.IndexOf("z-index:-1", StringComparison.OrdinalIgnoreCase) >= 0;
                p.FloatingObjects.Add(obj);
                return true;
            }

            private FloatingObject ReadEmbeddedObject(XElement element, Paragraph p, RunFormat fmt)
            {
                FloatingObject obj = new FloatingObject();
                obj.Kind = FloatingKind.Embedded;
                obj.Format = fmt.Clone();
                XElement ole = element.Descendants(OLE + "OLEObject").FirstOrDefault();
                obj.Text = ole != null ? ((string)ole.Attribute("ProgID") ?? "Objeto incrustado") : "Objeto incrustado";
                if (ole != null)
                {
                    obj.ProgId = (string)ole.Attribute("ProgID") ?? string.Empty;
                    obj.ShapeId = (string)ole.Attribute("ShapeID") ?? string.Empty;
                    string oleRid = (string)ole.Attribute(R + "id");
                    string oleTarget;
                    if (!string.IsNullOrEmpty(oleRid) && _rels.TryGetValue(oleRid, out oleTarget))
                    {
                        string normalized = oleTarget.Replace('\\', '/').TrimStart('/');
                        if (!normalized.StartsWith("word/", StringComparison.OrdinalIgnoreCase)) normalized = "word/" + normalized;
                        if (_zip.TryGetValue(normalized, out obj.EmbeddedData))
                        {
                            int slash = normalized.LastIndexOf('/');
                            obj.EmbeddedName = slash >= 0 ? normalized.Substring(slash + 1) : normalized;
                        }
                    }
                }
                XElement imageData = element.Descendants(VML + "imagedata").FirstOrDefault();
                if (imageData != null)
                {
                    string rid = (string)imageData.Attribute(R + "id");
                    string target;
                    if (!string.IsNullOrEmpty(rid) && _rels.TryGetValue(rid, out target))
                    {
                        string normalized = target.Replace('\\', '/');
                        if (normalized.StartsWith("./", StringComparison.Ordinal)) normalized = normalized.Substring(2);
                        byte[] data;
                        if (!_zip.TryGetValue("word/" + normalized, out data) && !_zip.TryGetValue(normalized, out data)) data = null;
                        if (data != null) obj.Image = new ImageInline(data, ExtensionOf(normalized), 2880, 1440, fmt.Clone());
                    }
                }
                if (obj.Image != null)
                {
                    obj.WidthTwips = obj.Image.WidthTwips;
                    obj.HeightTwips = obj.Image.HeightTwips;
                }
                obj.Wrap = TextWrapMode.InFront;
                return obj;
            }

            private static string ExtensionOf(string path)
            {
                int dot = path.LastIndexOf('.');
                if (dot < 0 || dot >= path.Length - 1) return "png";
                string ext = path.Substring(dot + 1).ToLowerInvariant();
                return ext == "jpg" ? "jpeg" : ext;
            }

            private static void ReadExtent(XElement extent, FloatingObject obj)
            {
                if (extent == null) return;
                long cx, cy;
                if (long.TryParse((string)extent.Attribute("cx"), NumberStyles.Integer, CultureInfo.InvariantCulture, out cx) && cx > 0)
                    obj.WidthTwips = (int)Math.Max(1, cx / 635);
                if (long.TryParse((string)extent.Attribute("cy"), NumberStyles.Integer, CultureInfo.InvariantCulture, out cy) && cy > 0)
                    obj.HeightTwips = (int)Math.Max(1, cy / 635);
            }

            private static int ReadPosition(XElement anchor, string name)
            {
                XElement position = anchor.Element(WP + name);
                if (position == null) return 0;
                XElement offset = position.Element(WP + "posOffset");
                long value;
                if (offset != null && long.TryParse((string)offset.Attribute("x") ?? (string)offset, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                    return (int)(value / 635);
                return 0;
            }

            private static TextWrapMode ReadWrap(XElement anchor)
            {
                if (anchor.Element(WP + "wrapNone") != null) return TextWrapMode.InFront;
                if (anchor.Element(WP + "wrapTopAndBottom") != null) return TextWrapMode.TopAndBottom;
                if (anchor.Element(WP + "wrapThrough") != null) return TextWrapMode.Through;
                if (anchor.Element(WP + "wrapTight") != null) return TextWrapMode.Tight;
                if (anchor.Element(WP + "wrapSquare") != null) return TextWrapMode.Square;
                return TextWrapMode.Square;
            }

            private static void ReadShapeFill(XElement drawing, FloatingObject obj)
            {
                XElement solid = drawing.Descendants(A + "solidFill").FirstOrDefault();
                if (solid != null)
                {
                    XElement color = solid.Element(A + "srgbClr");
                    if (color != null)
                    {
                        Color? parsed = ReadColor((string)color.Attribute("val"));
                        if (parsed.HasValue) obj.Fill = parsed;
                    }
                }
            }

            private static string CollectBlockText(XElement content)
            {
                if (content == null) return string.Empty;
                StringBuilder sb = new StringBuilder();
                foreach (XElement p in content.Descendants(W + "p"))
                {
                    if (sb.Length > 0) sb.Append('\n');
                    foreach (XElement t in p.Descendants(W + "t")) sb.Append(t.Value);
                }
                return sb.ToString();
            }

            private ImageInline ReadVmlImage(XElement pict)
            {
                XElement imageData = pict.Descendants(VML + "imagedata").FirstOrDefault();
                if (imageData == null) return null;
                string rid = (string)imageData.Attribute(R + "id");
                if (string.IsNullOrEmpty(rid)) return null;
                string target;
                if (!_rels.TryGetValue(rid, out target)) return null;
                string normalized = target.Replace('\\', '/');
                if (normalized.StartsWith("./", StringComparison.Ordinal)) normalized = normalized.Substring(2);
                if (!normalized.StartsWith("media/", StringComparison.OrdinalIgnoreCase))
                {
                    int idx = normalized.IndexOf("media/", StringComparison.OrdinalIgnoreCase);
                    if (idx >= 0) normalized = normalized.Substring(idx);
                }
                byte[] data;
                if (!_zip.TryGetValue("word/" + normalized, out data) && !_zip.TryGetValue(normalized, out data)) return null;

                int width = 1440, height = 1440;
                XElement shape = pict.Descendants(VML + "shape").FirstOrDefault();
                if (shape != null)
                {
                    string style = (string)shape.Attribute("style");
                    if (!string.IsNullOrEmpty(style)) ParseCssSize(style, ref width, ref height);
                }
                string ext = "png";
                int dot = normalized.LastIndexOf('.');
                if (dot >= 0 && dot < normalized.Length - 1)
                {
                    ext = normalized.Substring(dot + 1).ToLowerInvariant();
                    if (ext == "jpg") ext = "jpeg";
                }
                return new ImageInline(data, ext, width, height, new RunFormat());
            }

            private static void ParseCssSize(string style, ref int widthTwips, ref int heightTwips)
            {
                int w, h;
                if (TryCssLength(style, "width", out w)) widthTwips = w;
                if (TryCssLength(style, "height", out h)) heightTwips = h;
            }

            private static bool TryCssSignedLength(string style, string name, out int twips)
            {
                return TryCssLength(style, name, true, out twips);
            }

            private static bool TryCssLength(string style, string name, out int twips)
            {
                return TryCssLength(style, name, false, out twips);
            }

            private static bool TryCssLength(string style, string name, bool signed, out int twips)
            {
                twips = 0;
                int pos = style.IndexOf(name + ":", StringComparison.OrdinalIgnoreCase);
                if (pos < 0) return false;
                pos += name.Length + 1;
                while (pos < style.Length && char.IsWhiteSpace(style[pos])) pos++;
                int start = pos;
                while (pos < style.Length && (char.IsDigit(style[pos]) || style[pos] == '.' || style[pos] == '-')) pos++;
                if (pos == start) return false;
                double value;
                if (!double.TryParse(style.Substring(start, pos - start), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
                string unit = "px";
                int unitStart = pos;
                while (pos < style.Length && char.IsLetter(style[pos])) pos++;
                if (pos > unitStart) unit = style.Substring(unitStart, pos - unitStart).ToLowerInvariant();
                if (unit == "pt") twips = (int)Math.Round(value * 20);
                else if (unit == "in") twips = (int)Math.Round(value * 1440);
                else if (unit == "cm") twips = (int)Math.Round(value * 1440 / 2.54);
                else if (unit == "mm") twips = (int)Math.Round(value * 1440 / 25.4);
                else twips = (int)Math.Round(value * 15);
                return signed || twips > 0;
            }

            private ImageInline ReadDrawing(XElement drawing)
            {
                XElement blip = drawing.Descendants(A + "blip").FirstOrDefault();
                if (blip == null) return null;
                string rid = (string)blip.Attribute(R + "embed");
                if (string.IsNullOrEmpty(rid)) return null;
                string target;
                if (!_rels.TryGetValue(rid, out target)) return null;
                string normalized = target.Replace('\\', '/');
                if (normalized.StartsWith("./", StringComparison.Ordinal)) normalized = normalized.Substring(2);
                if (!normalized.StartsWith("media/", StringComparison.OrdinalIgnoreCase))
                {
                    int idx = normalized.IndexOf("media/", StringComparison.OrdinalIgnoreCase);
                    if (idx >= 0) normalized = normalized.Substring(idx);
                }
                string entryName = "word/" + normalized;
                byte[] data;
                if (!_zip.TryGetValue(entryName, out data))
                {
                    // try the target as-is
                    if (!_zip.TryGetValue(normalized, out data)) return null;
                }

                int width = 1440, height = 1440;
                XElement extent = drawing.Descendants(WP + "extent").FirstOrDefault();
                if (extent != null)
                {
                    long cx, cy;
                    if (long.TryParse((string)extent.Attribute("cx"), out cx) && cx > 0)
                        width = (int)Math.Max(1, cx / 635);
                    if (long.TryParse((string)extent.Attribute("cy"), out cy) && cy > 0)
                        height = (int)Math.Max(1, cy / 635);
                }

                string ext = "png";
                int dot = normalized.LastIndexOf('.');
                if (dot >= 0 && dot < normalized.Length - 1)
                {
                    ext = normalized.Substring(dot + 1).ToLowerInvariant();
                    if (ext == "jpg") ext = "jpeg";
                }
                return new ImageInline(data, ext, width, height, new RunFormat());
            }

            private void ReadTableCellRuns(XElement tc, TableCell cell)
            {
                if (cell == null) return;
                cell.Runs.Clear();
                foreach (XElement para in tc.Elements(W + "p"))
                {
                    Paragraph temp = new Paragraph();
                    ReadInlines(para, temp, _defaultRun.Clone(), null);
                    if (cell.Runs.Count > 0)
                        cell.Runs.Add(new TableCellRun("\n", _defaultRun.Clone()));
                    for (int i = 0; i < temp.Inlines.Count; i++)
                    {
                        TextInline run = temp.Inlines[i] as TextInline;
                        if (run == null) continue;
                        TableCellRun cellRun = new TableCellRun(run.Text, run.Format == null ? _defaultRun.Clone() : run.Format.Clone());
                        cellRun.LinkUrl = run.Format == null ? null : run.Format.LinkUrl;
                        cell.Runs.Add(cellRun);
                    }
                }
                if (cell.Runs.Count > 0)
                {
                    cell.SyncFromRuns();
                    bool anyLink = false;
                    for (int i = 0; i < cell.Runs.Count; i++)
                        if (!string.IsNullOrEmpty(cell.Runs[i].LinkUrl)) anyLink = true;
                    if (anyLink) cell.LinkUrl = null;
                }
            }
            private void ReadTable(XElement table, DocumentModel model)
            {
                List<List<TableCell>> cells = new List<List<TableCell>>();
                bool repeatHeader = false;
                int rowIndex = 0;
                int columns = 0;
                foreach (XElement row in table.Descendants(W + "tr"))
                {
                    if (row.Ancestors(W + "tbl").FirstOrDefault() != table) continue;
                    XElement rowProps = row.Element(W + "trPr");
                    if (rowIndex == 0 && rowProps != null && rowProps.Element(W + "tblHeader") != null)
                        repeatHeader = true;
                    List<TableCell> rowCells = new List<TableCell>();
                    int column = 0;
                    foreach (XElement tc in row.Elements(W + "tc"))
                    {
                        StringBuilder cellText = new StringBuilder();
                        foreach (XElement para in tc.Descendants(W + "p"))
                        {
                            if (cellText.Length > 0) cellText.Append(' ');
                            cellText.Append(CollectText(para));
                        }
                        XElement tcPr = tc.Element(W + "tcPr");
                        int gridSpan = 1;
                        if (tcPr != null)
                        {
                            XElement span = tcPr.Element(W + "gridSpan");
                            if (span != null) gridSpan = Math.Max(1, ReadIntAttr(span, "val"));
                        }
                        XElement vMerge = tcPr != null ? tcPr.Element(W + "vMerge") : null;
                        bool verticalContinue = vMerge != null &&
                            !string.Equals((string)vMerge.Attribute(W + "val"), "restart", StringComparison.OrdinalIgnoreCase);
                        if (verticalContinue && cells.Count > 0)
                        {
                            TableCell owner = FindOwner(cells, cells.Count - 1, column);
                            if (owner != null)
                            {
                                owner.RowSpan = Math.Max(owner.RowSpan, owner.RowSpan + 1);
                                for (int i = 0; i < gridSpan; i++)
                                {
                                    while (rowCells.Count <= column) rowCells.Add(new TableCell());
                                    rowCells[column] = new TableCell { Covered = true };
                                    column++;
                                }
                                continue;
                            }
                        }
                        TableCell cell = new TableCell();
                        cell.Text = cellText.ToString();
                        ReadTableCellRuns(tc, cell);
                        cell.ColSpan = gridSpan;
                        XElement cellLink = tc.Descendants(W + "hyperlink").FirstOrDefault();
                        if (cellLink != null)
                        {
                            string cellRid = (string)cellLink.Attribute(R + "id");
                            if (!string.IsNullOrEmpty(cellRid) && _rels.ContainsKey(cellRid))
                                cell.LinkUrl = _rels[cellRid];
                        }
                        XElement firstRun = null;
                        foreach (XElement run in tc.Descendants(W + "r"))
                        {
                            firstRun = run;
                            break;
                        }
                        if (firstRun != null)
                        {
                            XElement runProps = firstRun.Element(W + "rPr");
                            if (runProps != null)
                            {
                                RunFormat cellFormat = new RunFormat("Calibri", 21);
                                ApplyRunProperties(cellFormat, runProps);
                                cell.Format = cellFormat;
                                cell.Bold = cellFormat.Bold;
                                cell.Italic = cellFormat.Italic;
                            }
                        }
                        if (tcPr != null)
                        {
                            XElement shd = tcPr.Element(W + "shd");
                            if (shd != null) cell.Shading = ReadColor((string)shd.Attribute(W + "fill"));
                            XElement tcBorders = tcPr.Element(W + "tcBorders");
                            if (tcBorders != null) ReadBorders(tcBorders, cell);
                        }
                        while (rowCells.Count < column) rowCells.Add(new TableCell { Covered = true });
                        rowCells.Add(cell);
                        for (int i = 1; i < gridSpan; i++) rowCells.Add(new TableCell { Covered = true });
                        column += gridSpan;
                    }
                    if (rowCells.Count > columns) columns = rowCells.Count;
                    cells.Add(rowCells);
                    rowIndex++;
                }
                if (cells.Count == 0) return;

                TableCell[][] grid = new TableCell[cells.Count][];
                for (int r = 0; r < cells.Count; r++)
                {
                    grid[r] = new TableCell[columns];
                    for (int c = 0; c < columns; c++)
                        grid[r][c] = c < cells[r].Count && cells[r][c] != null ? cells[r][c] : new TableCell { Covered = true };
                }
                Paragraph p = new Paragraph();
                CopyParagraphFormat(_defaultPara, p);
                XElement tableProperties = table.Element(W + "tblPr");
                if (tableProperties != null)
                {
                    XElement style = tableProperties.Element(W + "tblStyle");
                    if (style != null) p.TableStyleName = (string)style.Attribute(W + "val") ?? string.Empty;
                    XElement tableBorders = tableProperties.Element(W + "tblBorders");
                    if (tableBorders != null) ReadBorders(tableBorders, null, p);
                    XElement tableShd = tableProperties.Element(W + "shd");
                    if (tableShd != null) p.TableShading = ReadColor((string)tableShd.Attribute(W + "fill"));
                }
                if (string.IsNullOrEmpty(p.TableStyleName)) p.TableStyleName = "TableGrid";
                ApplyKnownTableStyle(p, grid);
                p.TableCells = grid;
                p.RepeatTableHeader = repeatHeader;
                p.SyncTableRows();
                int width = 0;
                XElement tblW = table.Element(W + "tblPr") != null
                    ? table.Element(W + "tblPr").Element(W + "tblW") : null;
                if (tblW != null) width = ReadIntAttr(tblW, "w");
                p.TableWidthTwips = width > 0 ? width : 0;
                model.Paragraphs.Add(p);
            }

            private static Color? ReadColor(string value)
            {
                if (string.IsNullOrEmpty(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) return null;
                try
                {
                    if (value.Length == 6) return Color.FromArgb(
                        Convert.ToInt32(value.Substring(0, 2), 16),
                        Convert.ToInt32(value.Substring(2, 2), 16),
                        Convert.ToInt32(value.Substring(4, 2), 16));
                }
                catch { }
                return null;
            }

            private static void ReadBorders(XElement borders, TableCell cell)
            {
                if (borders == null || cell == null) return;
                foreach (XElement border in borders.Elements())
                {
                    string color = (string)border.Attribute(W + "color");
                    int size = ReadIntAttr(border, "sz");
                    Color? parsed = ReadColor(color);
                    if (parsed.HasValue) { cell.BorderColor = parsed; cell.BorderSize = Math.Max(1, size / 8); return; }
                }
            }

            private static void ReadBorders(XElement borders, TableCell cell, Paragraph table)
            {
                if (borders == null || table == null) return;
                foreach (XElement border in borders.Elements())
                {
                    Color? parsed = ReadColor((string)border.Attribute(W + "color"));
                    int size = ReadIntAttr(border, "sz");
                    if (parsed.HasValue)
                    {
                        table.TableBorderColor = parsed;
                        table.TableBorderSize = Math.Max(1, size / 8);
                        return;
                    }
                }
            }

            private static void ApplyKnownTableStyle(Paragraph table, TableCell[][] cells)
            {
                string style = table.TableStyleName ?? string.Empty;
                bool grid = style.IndexOf("Grid", StringComparison.OrdinalIgnoreCase) >= 0
                    || style.IndexOf("Cuadr", StringComparison.OrdinalIgnoreCase) >= 0;
                if (grid && !table.TableBorderColor.HasValue)
                {
                    table.TableBorderColor = Color.FromArgb(125, 132, 140);
                    table.TableBorderSize = 1;
                }
                bool banded = style.IndexOf("Light", StringComparison.OrdinalIgnoreCase) >= 0
                    || style.IndexOf("Medium", StringComparison.OrdinalIgnoreCase) >= 0;
                if (cells == null) return;
                for (int r = 0; r < cells.Length; r++)
                {
                    for (int c = 0; c < cells[r].Length; c++)
                    {
                        TableCell cell = cells[r][c];
                        if (cell == null || cell.Covered) continue;
                        if (r == 0 && banded)
                        {
                            cell.Bold = true;
                            if (!cell.Shading.HasValue) cell.Shading = Color.FromArgb(217, 226, 243);
                        }
                        else if (banded && r % 2 == 0 && !cell.Shading.HasValue)
                            cell.Shading = Color.FromArgb(242, 242, 242);
                    }
                }
            }

            private static TableCell FindOwner(List<List<TableCell>> cells, int row, int column)
            {
                for (int r = row; r >= 0; r--)
                {
                    if (r >= cells.Count) continue;
                    for (int c = 0; c < cells[r].Count; c++)
                    {
                        TableCell cell = cells[r][c];
                        if (cell == null || cell.Covered) continue;
                        if (row >= r && row < r + Math.Max(1, cell.RowSpan) &&
                            column >= c && column < c + Math.Max(1, cell.ColSpan)) return cell;
                    }
                }
                return null;
            }

            private static string CollectText(XElement p)
            {
                StringBuilder sb = new StringBuilder();
                foreach (XElement t in p.Descendants(W + "t")) sb.Append(t.Value);
                foreach (XElement tab in p.Descendants(W + "tab")) sb.Append('\t');
                return sb.ToString();
            }

            private void ReadSectionInto(XElement sectPr, DocumentSection section)
            {
                XElement pgSz = sectPr.Element(W + "pgSz");
                if (pgSz != null)
                {
                    int w = ReadIntAttr(pgSz, "w");
                    int h = ReadIntAttr(pgSz, "h");
                    if (w > 0) section.PageWidthTwips = w;
                    if (h > 0) section.PageHeightTwips = h;
                }
                XElement pgMar = sectPr.Element(W + "pgMar");
                if (pgMar != null)
                {
                    int v;
                    v = ReadIntAttr(pgMar, "left"); if (v > 0) section.MarginLeftTwips = v;
                    v = ReadIntAttr(pgMar, "right"); if (v > 0) section.MarginRightTwips = v;
                    v = ReadIntAttr(pgMar, "top"); if (v > 0) section.MarginTopTwips = v;
                    v = ReadIntAttr(pgMar, "bottom"); if (v > 0) section.MarginBottomTwips = v;
                    if (pgMar.Attribute(W + "header") != null)
                    {
                        v = ReadIntAttr(pgMar, "header");
                        if (v >= 0) section.HeaderDistanceTwips = v;
                    }
                    if (pgMar.Attribute(W + "footer") != null)
                    {
                        v = ReadIntAttr(pgMar, "footer");
                        if (v >= 0) section.FooterDistanceTwips = v;
                    }
                }
                XElement cols = sectPr.Element(W + "cols");
                if (cols != null)
                {
                    int count = ReadIntAttr(cols, "num");
                    if (count > 1) section.ColumnCount = Math.Min(6, count);
                    int space = ReadIntAttr(cols, "space");
                    if (space > 0) section.ColumnSpacingTwips = space;
                    section.ColumnsSeparator = string.Equals((string)cols.Attribute(W + "sep"), "1", StringComparison.OrdinalIgnoreCase);
                }
                XElement type = sectPr.Element(W + "type");
                if (type != null)
                {
                    string value = (string)type.Attribute(W + "val");
                    if (string.Equals(value, "continuous", StringComparison.OrdinalIgnoreCase)) section.BreakType = SectionBreakType.Continuous;
                    else if (string.Equals(value, "evenPage", StringComparison.OrdinalIgnoreCase)) section.BreakType = SectionBreakType.EvenPage;
                    else if (string.Equals(value, "oddPage", StringComparison.OrdinalIgnoreCase)) section.BreakType = SectionBreakType.OddPage;
                }
                XElement titlePg = sectPr.Element(W + "titlePg");
                section.TitlePage = titlePg != null;
                section.HeaderText = ReadHeaderFooter(sectPr, "headerReference", "default");
                section.FooterText = ReadHeaderFooter(sectPr, "footerReference", "default");
                section.FirstPageHeaderText = ReadHeaderFooter(sectPr, "headerReference", "first");
                section.FirstPageFooterText = ReadHeaderFooter(sectPr, "footerReference", "first");
                section.EvenPageHeaderText = ReadHeaderFooter(sectPr, "headerReference", "even");
                section.EvenPageFooterText = ReadHeaderFooter(sectPr, "footerReference", "even");
                section.DifferentOddEven = HasHeaderFooterReference(sectPr, "headerReference", "even")
                    || HasHeaderFooterReference(sectPr, "footerReference", "even");
                ReadHeaderFooterImages(sectPr, "headerReference", "default", section.HeaderImages);
                ReadHeaderFooterImages(sectPr, "footerReference", "default", section.FooterImages);
                ReadHeaderFooterImages(sectPr, "headerReference", "first", section.FirstPageHeaderImages);
                ReadHeaderFooterImages(sectPr, "footerReference", "first", section.FirstPageFooterImages);
                ReadHeaderFooterImages(sectPr, "headerReference", "even", section.EvenPageHeaderImages);
                ReadHeaderFooterImages(sectPr, "footerReference", "even", section.EvenPageFooterImages);
                ReadHeaderFooterFormat(sectPr, "headerReference", "default", true, section);
                ReadHeaderFooterFormat(sectPr, "footerReference", "default", false, section);
                ReadHeaderFooterFormat(sectPr, "headerReference", "first", true, section);
                ReadHeaderFooterFormat(sectPr, "footerReference", "first", false, section);
                ReadHeaderFooterFormat(sectPr, "headerReference", "even", true, section);
                ReadHeaderFooterFormat(sectPr, "footerReference", "even", false, section);
            }

            private static bool HasHeaderFooterReference(XElement sectPr, string referenceName, string type)
            {
                foreach (XElement item in sectPr.Elements(W + referenceName))
                {
                    string itemType = (string)item.Attribute(W + "type");
                    if (string.IsNullOrEmpty(itemType)) itemType = "default";
                    if (string.Equals(itemType, type, StringComparison.OrdinalIgnoreCase)) return true;
                }
                return false;
            }

            private void ReadHeaderFooterImages(XElement sectPr, string referenceName, string type, List<ImageInline> target)
            {
                XElement reference = null;
                foreach (XElement item in sectPr.Elements(W + referenceName))
                {
                    string itemType = (string)item.Attribute(W + "type");
                    if (string.IsNullOrEmpty(itemType)) itemType = "default";
                    if (string.Equals(itemType, type, StringComparison.OrdinalIgnoreCase))
                    {
                        reference = item;
                        break;
                    }
                }
                if (reference == null) return;
                string rid = (string)reference.Attribute(R + "id");
                string path;
                if (string.IsNullOrEmpty(rid) || !_rels.TryGetValue(rid, out path)) return;
                string normalized = path.Replace('\\', '/').TrimStart('/');
                if (!normalized.StartsWith("word/", StringComparison.OrdinalIgnoreCase)) normalized = "word/" + normalized;
                byte[] data;
                if (!_zip.TryGetValue(normalized, out data)) return;
                try
                {
                    XDocument doc = LoadXml(data);
                    foreach (XElement drawing in doc.Descendants(W + "drawing"))
                    {
                        ImageInline image = ReadDrawing(drawing);
                        if (image != null) target.Add(image);
                    }
                    foreach (XElement pict in doc.Descendants(W + "pict"))
                    {
                        ImageInline image = ReadVmlImage(pict);
                        if (image != null) target.Add(image);
                    }
                }
                catch { }
            }

            private void ReadHeaderFooterFormat(XElement sectPr, string referenceName, string type, bool header,
                DocumentSection section)
            {
                XElement reference = null;
                foreach (XElement item in sectPr.Elements(W + referenceName))
                {
                    string itemType = (string)item.Attribute(W + "type");
                    if (string.IsNullOrEmpty(itemType)) itemType = "default";
                    if (string.Equals(itemType, type, StringComparison.OrdinalIgnoreCase))
                    {
                        reference = item;
                        break;
                    }
                }
                if (reference == null) return;
                string rid = (string)reference.Attribute(R + "id");
                string path;
                if (string.IsNullOrEmpty(rid) || !_rels.TryGetValue(rid, out path)) return;
                string normalized = path.Replace('\\', '/').TrimStart('/');
                if (!normalized.StartsWith("word/", StringComparison.OrdinalIgnoreCase)) normalized = "word/" + normalized;
                byte[] data;
                if (!_zip.TryGetValue(normalized, out data)) return;
                try
                {
                    XDocument doc = LoadXml(data);
                    RunFormat format = new RunFormat("Calibri", 18);
                    ParagraphAlignment alignment = ParagraphAlignment.Left;
                    foreach (XElement p in doc.Descendants(W + "p"))
                    {
                        XElement pPr = p.Element(W + "pPr");
                        if (pPr != null)
                        {
                            XElement jc = pPr.Element(W + "jc");
                            if (jc != null)
                            {
                                string value = (string)jc.Attribute(W + "val");
                                if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) alignment = ParagraphAlignment.Center;
                                else if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)) alignment = ParagraphAlignment.Right;
                                else if (string.Equals(value, "both", StringComparison.OrdinalIgnoreCase)) alignment = ParagraphAlignment.Justify;
                            }
                        }
                        XElement run = p.Element(W + "r");
                        if (run == null) continue;
                        XElement rPr = run.Element(W + "rPr");
                        if (rPr != null) ApplyRunProperties(format, rPr);
                        break;
                    }
                    if (string.Equals(type, "first", StringComparison.OrdinalIgnoreCase))
                    {
                        if (header) { section.FirstPageHeaderFormat = format; section.FirstPageHeaderAlignment = alignment; }
                        else { section.FirstPageFooterFormat = format; section.FirstPageFooterAlignment = alignment; }
                    }
                    else if (string.Equals(type, "even", StringComparison.OrdinalIgnoreCase))
                    {
                        if (header) { section.EvenPageHeaderFormat = format; section.EvenPageHeaderAlignment = alignment; }
                        else { section.EvenPageFooterFormat = format; section.EvenPageFooterAlignment = alignment; }
                    }
                    else if (header)
                    {
                        section.HeaderFormat = format;
                        section.HeaderAlignment = alignment;
                    }
                    else
                    {
                        section.FooterFormat = format;
                        section.FooterAlignment = alignment;
                    }
                }
                catch
                {
                }
            }

            private string ReadHeaderFooter(XElement sectPr, string referenceName, string type)
            {
                XElement reference = null;
                foreach (XElement item in sectPr.Elements(W + referenceName))
                {
                    string itemType = (string)item.Attribute(W + "type");
                    if (string.IsNullOrEmpty(itemType)) itemType = "default";
                    if (string.Equals(itemType, type, StringComparison.OrdinalIgnoreCase))
                    {
                        reference = item;
                        break;
                    }
                }
                if (reference == null) return string.Empty;
                string rid = (string)reference.Attribute(R + "id");
                if (string.IsNullOrEmpty(rid)) return string.Empty;
                string target;
                if (!_rels.TryGetValue(rid, out target)) return string.Empty;
                string normalized = target.Replace('\\', '/').TrimStart('/');
                if (!normalized.StartsWith("word/", StringComparison.OrdinalIgnoreCase)) normalized = "word/" + normalized;
                byte[] data;
                if (!_zip.TryGetValue(normalized, out data)) return string.Empty;
                try
                {
                    XDocument doc = LoadXml(data);
                    List<string> lines = new List<string>();
                    foreach (XElement p in doc.Descendants(W + "p"))
                    {
                        lines.Add(CollectHeaderFooterParagraph(p));
                    }
                    return string.Join("\n", lines.ToArray()).TrimEnd('\n');
                }
                catch
                {
                    return string.Empty;
                }
            }

            private static string CollectHeaderFooterParagraph(XElement paragraph)
            {
                StringBuilder text = new StringBuilder();
                int fieldDepth = 0;
                bool showingResult = false;
                string instruction = string.Empty;
                string placeholder = null;
                foreach (XElement node in paragraph.Descendants())
                {
                    if (node.Name == W + "fldChar")
                    {
                        string type = (string)node.Attribute(W + "fldCharType");
                        if (string.Equals(type, "begin", StringComparison.OrdinalIgnoreCase))
                        {
                            fieldDepth++;
                            instruction = string.Empty;
                            placeholder = null;
                            showingResult = false;
                        }
                        else if (string.Equals(type, "separate", StringComparison.OrdinalIgnoreCase))
                        {
                            showingResult = true;
                            string upper = instruction.ToUpperInvariant();
                            if (upper.IndexOf("NUMPAGES", StringComparison.Ordinal) >= 0) placeholder = "{NUMPAGES}";
                            else if (upper.IndexOf("PAGE", StringComparison.Ordinal) >= 0) placeholder = "{PAGE}";
                        }
                        else if (string.Equals(type, "end", StringComparison.OrdinalIgnoreCase))
                        {
                            if (placeholder != null) text.Append(placeholder);
                            if (fieldDepth > 0) fieldDepth--;
                            showingResult = false;
                            placeholder = null;
                        }
                        continue;
                    }
                    if (node.Name == W + "instrText")
                    {
                        if (fieldDepth > 0 && !showingResult) instruction += node.Value;
                        continue;
                    }
                    if (fieldDepth > 0 && (!showingResult || placeholder != null)) continue;
                    if (node.Name == W + "t") text.Append(node.Value);
                    else if (node.Name == W + "tab") text.Append('\t');
                    else if (node.Name == W + "br" || node.Name == W + "cr") text.Append('\n');
                }
                return text.ToString();
            }

            private void ApplyCharacterStyle(RunFormat target, string id, HashSet<string> seen)
            {
                if (string.IsNullOrEmpty(id) || seen.Contains(id)) return;
                seen.Add(id);
                StyleDef def;
                if (!_styles.TryGetValue(id, out def)) return;
                if (!string.IsNullOrEmpty(def.BasedOn))
                    ApplyCharacterStyle(target, def.BasedOn, seen);
                if (def.RPr != null) ApplyRunProperties(target, def.RPr);
            }

            private Paragraph GetParagraphStyle(string id)
            {
                Paragraph cached;
                if (_styleParaCache.TryGetValue(id, out cached)) return cached;
                if (_styleParaCache.ContainsKey(id)) return null;
                _styleParaCache[id] = null;
                StyleDef def;
                if (!_styles.TryGetValue(id, out def)) return null;
                Paragraph p = new Paragraph();
                CopyParagraphFormat(_defaultPara, p);
                if (!string.IsNullOrEmpty(def.BasedOn) && _styles.ContainsKey(def.BasedOn))
                {
                    Paragraph parent = GetParagraphStyle(def.BasedOn);
                    if (parent != null) CopyParagraphFormat(parent, p);
                }
                if (def.PPr != null) ApplyParagraphProperties(p, def.PPr);
                _styleParaCache[id] = p;
                return p;
            }

            private RunFormat GetRunStyle(string id)
            {
                RunFormat cached;
                if (_styleRunCache.TryGetValue(id, out cached)) return cached;
                if (_styleRunCache.ContainsKey(id)) return null;
                _styleRunCache[id] = null;
                StyleDef def;
                if (!_styles.TryGetValue(id, out def)) return null;
                RunFormat f = _defaultRun.Clone();
                if (!string.IsNullOrEmpty(def.BasedOn) && _styles.ContainsKey(def.BasedOn))
                {
                    RunFormat parent = GetRunStyle(def.BasedOn);
                    if (parent != null) f = parent.Clone();
                }
                if (def.RPr != null) ApplyRunProperties(f, def.RPr);
                _styleRunCache[id] = f;
                return f;
            }

            private static void CopyParagraphFormat(Paragraph from, Paragraph to)
            {
                to.Alignment = from.Alignment;
                to.LeftIndentTwips = from.LeftIndentTwips;
                to.RightIndentTwips = from.RightIndentTwips;
                to.FirstLineIndentTwips = from.FirstLineIndentTwips;
                to.SpaceBeforeTwips = from.SpaceBeforeTwips;
                to.SpaceAfterTwips = from.SpaceAfterTwips;
                to.LineSpacing = from.LineSpacing;
                to.TabStopsTwips = (int[])from.TabStopsTwips.Clone();
            }

            private static void ApplyParagraphProperties(Paragraph p, XElement pPr)
            {
                if (pPr == null) return;
                XElement jc = pPr.Element(W + "jc");
                if (jc != null)
                {
                    string val = (string)jc.Attribute(W + "val");
                    if (val == "center") p.Alignment = ParagraphAlignment.Center;
                    else if (val == "right" || val == "end") p.Alignment = ParagraphAlignment.Right;
                    else if (val == "both" || val == "distribute") p.Alignment = ParagraphAlignment.Justify;
                    else if (val == "left" || val == "start") p.Alignment = ParagraphAlignment.Left;
                }
                XElement ind = pPr.Element(W + "ind");
                if (ind != null)
                {
                    int v = ReadIntAttr(ind, "left", "start");
                    if (v != 0) p.LeftIndentTwips = v;
                    v = ReadIntAttr(ind, "right", "end");
                    if (v != 0) p.RightIndentTwips = v;
                    v = ReadIntAttr(ind, "firstLine");
                    if (v != 0) p.FirstLineIndentTwips = v;
                    v = ReadIntAttr(ind, "hanging");
                    if (v != 0) p.FirstLineIndentTwips = -v;
                }
                XElement spacing = pPr.Element(W + "spacing");
                if (spacing != null)
                {
                    int before = ReadIntAttr(spacing, "before");
                    int after = ReadIntAttr(spacing, "after");
                    if (before != 0) p.SpaceBeforeTwips = before;
                    if (after != 0) p.SpaceAfterTwips = after;
                    string rule = (string)spacing.Attribute(W + "lineRule");
                    int line = ReadIntAttr(spacing, "line");
                    if (line > 0 && (rule == null || rule == "auto"))
                    {
                        if (line >= 460) p.LineSpacing = 2;
                        else if (line >= 320) p.LineSpacing = 1;
                        else p.LineSpacing = 0;
                    }
                }
                p.PageBreakBefore = ReadOnOff(pPr.Element(W + "pageBreakBefore"), p.PageBreakBefore);
                p.KeepLinesTogether = ReadOnOff(pPr.Element(W + "keepLines"), p.KeepLinesTogether);
                p.KeepWithNext = ReadOnOff(pPr.Element(W + "keepNext"), p.KeepWithNext);
                p.WidowControl = ReadOnOff(pPr.Element(W + "widowControl"), p.WidowControl);
                p.NoSpaceBetweenSameStyle = ReadOnOff(pPr.Element(W + "contextualSpacing"), p.NoSpaceBetweenSameStyle);
                XElement outline = pPr.Element(W + "outlineLvl");
                if (outline != null) p.OutlineLevel = Math.Max(0, Math.Min(8, ReadIntAttr(outline, "val")));
                XElement tabs = pPr.Element(W + "tabs");
                if (tabs != null)
                {
                    List<int> positions = new List<int>();
                    foreach (XElement tab in tabs.Elements(W + "tab"))
                    {
                        int pos = ReadIntAttr(tab, "pos");
                        if (pos > 0) positions.Add(pos);
                    }
                    if (positions.Count > 0) p.TabStopsTwips = positions.ToArray();
                }
            }

            private static void ApplyRunProperties(RunFormat f, XElement rPr)
            {
                if (rPr == null) return;
                XElement fonts = rPr.Element(W + "rFonts");
                if (fonts != null)
                {
                    string name = (string)fonts.Attribute(W + "ascii");
                    if (string.IsNullOrEmpty(name)) name = (string)fonts.Attribute(W + "hAnsi");
                    if (string.IsNullOrEmpty(name)) name = (string)fonts.Attribute(W + "cs");
                    if (string.IsNullOrEmpty(name)) name = (string)fonts.Attribute(W + "eastAsia");
                    if (!string.IsNullOrEmpty(name)) f.FontFamily = name;
                }
                f.Bold = ReadOnOff(rPr.Element(W + "b"), f.Bold);
                f.Italic = ReadOnOff(rPr.Element(W + "i"), f.Italic);
                f.Strikeout = ReadOnOff(rPr.Element(W + "strike"), ReadOnOff(rPr.Element(W + "dstrike"), f.Strikeout));
                XElement u = rPr.Element(W + "u");
                if (u != null)
                {
                    string val = (string)u.Attribute(W + "val");
                    f.Underline = val != "none" && val != "0" && val != "false";
                }
                XElement color = rPr.Element(W + "color");
                if (color != null)
                {
                    string val = (string)color.Attribute(W + "val");
                    Color c;
                    if (val == "auto") f.ForeColor = null;
                    else if (TryHexColor(val, out c)) f.ForeColor = c;
                }
                XElement highlight = rPr.Element(W + "highlight");
                if (highlight != null)
                {
                    Color c;
                    if (TryHighlight((string)highlight.Attribute(W + "val"), out c)) f.BackColor = c;
                }
                XElement shd = rPr.Element(W + "shd");
                if (shd != null)
                {
                    string fill = (string)shd.Attribute(W + "fill");
                    Color c;
                    if (!string.IsNullOrEmpty(fill) && fill != "auto" && TryHexColor(fill, out c)) f.BackColor = c;
                }
                XElement sz = rPr.Element(W + "sz");
                if (sz != null)
                {
                    int v;
                    if (int.TryParse((string)sz.Attribute(W + "val"), out v) && v > 0) f.SizeHalfPoints = v;
                }
                XElement va = rPr.Element(W + "vertAlign");
                if (va != null)
                {
                    string val = (string)va.Attribute(W + "val");
                    if (val == "superscript") f.Script = 1;
                    else if (val == "subscript") f.Script = -1;
                    else f.Script = 0;
                }
            }

            private static bool ReadOnOff(XElement e, bool current)
            {
                if (e == null) return current;
                string val = (string)e.Attribute(W + "val");
                if (val == null) return true;
                return val != "0" && val != "false" && val != "off";
            }

            private static int ReadIntAttr(XElement e, params string[] names)
            {
                for (int i = 0; i < names.Length; i++)
                {
                    XAttribute a = e.Attribute(W + names[i]);
                    if (a != null)
                    {
                        int v;
                        if (int.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out v)) return v;
                    }
                }
                return 0;
            }

            private static int ReadIntElement(XElement parent, string name)
            {
                XElement e = parent.Element(W + name);
                if (e == null) return 0;
                int v;
                if (int.TryParse((string)e.Attribute(W + "val"), out v)) return v;
                return 0;
            }

            private static bool TryHexColor(string val, out Color color)
            {
                color = Color.Black;
                if (string.IsNullOrEmpty(val)) return false;
                val = val.TrimStart('#');
                if (val.Length != 6 && val.Length != 8) return false;
                int r, g, b;
                if (!int.TryParse(val.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r)) return false;
                if (!int.TryParse(val.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g)) return false;
                if (!int.TryParse(val.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b)) return false;
                color = Color.FromArgb(r, g, b);
                return true;
            }

            private static bool TryHighlight(string name, out Color color)
            {
                color = Color.Yellow;
                if (string.IsNullOrEmpty(name)) return false;
                switch (name.ToLowerInvariant())
                {
                    case "yellow": color = Color.Yellow; return true;
                    case "green": color = Color.Lime; return true;
                    case "cyan": color = Color.Cyan; return true;
                    case "magenta": color = Color.Magenta; return true;
                    case "blue": color = Color.Blue; return true;
                    case "red": color = Color.Red; return true;
                    case "darkblue": color = Color.DarkBlue; return true;
                    case "darkcyan": color = Color.DarkCyan; return true;
                    case "darkgreen": color = Color.DarkGreen; return true;
                    case "darkmagenta": color = Color.DarkMagenta; return true;
                    case "darkred": color = Color.DarkRed; return true;
                    case "darkyellow": color = Color.Olive; return true;
                    case "darkgray": color = Color.DarkGray; return true;
                    case "lightgray": color = Color.LightGray; return true;
                    case "black": color = Color.Black; return true;
                    case "white": color = Color.White; return true;
                    case "none": return false;
                }
                return false;
            }

            private static string ToLetters(int n, bool upper)
            {
                if (n <= 0) n = 1;
                StringBuilder sb = new StringBuilder();
                while (n > 0)
                {
                    n--;
                    sb.Insert(0, (char)('a' + (n % 26)));
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
    }
}
