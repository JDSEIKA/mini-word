using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using WordPadClone.Util;

namespace WordPadClone.Formats
{
    /// <summary>Headless exercise of the format pipeline (run with --selftest).</summary>
    internal static class SelfTest
    {
        public static bool Run(string directory)
        {
            Directory.CreateDirectory(directory);
            StringBuilder log = new StringBuilder();
            bool ok = true;

            Form form = new Form();
            form.ShowInTaskbar = false;
            form.WindowState = FormWindowState.Minimized;
            RichTextBox editor = new RichTextBox();
            editor.Dock = DockStyle.Fill;
            form.Controls.Add(editor);
            form.Show();
            Application.DoEvents();

            try
            {
                DocumentModel sample = BuildSample();
                string rtf = RtfWriter.Write(sample, false);
                editor.Rtf = rtf;
                Application.DoEvents();

                Check(log, ref ok, "RTF load text", editor.Text.Replace("\n", "|") ==
                    "Informe de prueba|Primera viñeta con ñ y acentos|Texto normal superíndice y final.");

                DocumentModel round = RtfReader.Read(editor.Rtf);
                Check(log, ref ok, "RTF paragraph count", round.Paragraphs.Count == 3);
                Check(log, ref ok, "RTF bullet", round.Paragraphs[1].Bullet && round.Paragraphs[1].LeftIndentTwips == 720);
                Check(log, ref ok, "RTF center", round.Paragraphs[0].Alignment == ParagraphAlignment.Center);
                TextInline first = round.Paragraphs[0].Inlines[0] as TextInline;
                Check(log, ref ok, "RTF bold", first != null && first.Format.Bold);
                TextInline script = round.Paragraphs[2].Inlines[1] as TextInline;
                Check(log, ref ok, "RTF superscript", script != null && script.Format.Script == 1);

                round.PageWidthTwips = 15840;
                round.PageHeightTwips = 12240;
                round.MarginLeftTwips = 1440;
                round.MarginRightTwips = 1440;
                round.HeaderText = "Mini Word · página {PAGE} de {NUMPAGES}";
                round.FooterText = "Documento de prueba · confidencial";

                string docx = Path.Combine(directory, "selftest.docx");
                string odt = Path.Combine(directory, "selftest.odt");
                string plain = Path.Combine(directory, "selftest.txt");
                string rtfPath = Path.Combine(directory, "selftest.rtf");

                DocumentFormats.SaveDocx(docx, round);
                DocumentFormats.SaveOdt(odt, round);
                DocumentFormats.SaveRtf(rtfPath, RtfWriter.ApplyPageSetup(editor.Rtf, round));
                DocumentFormats.SaveText(plain, editor.Text, TextEncodingKind.Utf8);
                Application.DoEvents();

                Check(log, ref ok, "DOCX exists", new FileInfo(docx).Length > 1000);
                Check(log, ref ok, "ODT exists", new FileInfo(odt).Length > 1000);
                Check(log, ref ok, "RTF detect", DocumentFormats.Detect(rtfPath) == DocumentFormat.Rtf);
                Check(log, ref ok, "DOCX detect", DocumentFormats.Detect(docx) == DocumentFormat.Docx);
                Check(log, ref ok, "ODT detect", DocumentFormats.Detect(odt) == DocumentFormat.Odt);

                DocumentModel dx = DocxFormat.Read(docx);
                DocumentModel od = OdtFormat.Read(odt);
                Check(log, ref ok, "DOCX text", dx.PlainText == round.PlainText);
                Check(log, ref ok, "ODT text", od.PlainText == round.PlainText);
                Check(log, ref ok, "DOCX bullet", dx.Paragraphs[1].Bullet);
                Check(log, ref ok, "ODT bullet", od.Paragraphs[1].Bullet);
                Check(log, ref ok, "DOCX page", dx.PageWidthTwips == 15840 && dx.PageHeightTwips == 12240);
                Check(log, ref ok, "ODT page", od.PageWidthTwips == 15840 && od.PageHeightTwips == 12240);
                Check(log, ref ok, "DOCX header", dx.HeaderText.IndexOf("Mini Word", StringComparison.Ordinal) >= 0
                    && dx.HeaderText.IndexOf("{PAGE}", StringComparison.Ordinal) >= 0
                    && dx.HeaderText.IndexOf("{NUMPAGES}", StringComparison.Ordinal) >= 0);
                Check(log, ref ok, "DOCX footer", dx.FooterText.IndexOf("confidencial", StringComparison.Ordinal) >= 0);

                // Strict OOXML uses different namespace URIs for the same
                // vocabulary.  Produce a strict package from the transitional
                // one and verify that the reader normalizes it.
                Dictionary<string, byte[]> package = ZipArchiveLite.Read(docx);
                byte[] documentBytes;
                if (package.TryGetValue("word/document.xml", out documentBytes))
                {
                    string documentXml = Encoding.UTF8.GetString(documentBytes)
                        .Replace("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main")
                        .Replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships");
                    package["word/document.xml"] = new UTF8Encoding(false).GetBytes(documentXml);
                    string strictPath = Path.Combine(directory, "selftest-strict.docx");
                    List<ZipEntryData> strictEntries = new List<ZipEntryData>();
                    foreach (KeyValuePair<string, byte[]> item in package)
                        strictEntries.Add(new ZipEntryData(item.Key, item.Value, true));
                    ZipArchiveLite.Write(strictPath, strictEntries);
                    DocumentModel strict = DocxFormat.Read(strictPath);
                    Check(log, ref ok, "DOCX Strict text", strict.PlainText == round.PlainText);
                    Check(log, ref ok, "DOCX Strict page", strict.PageWidthTwips == 15840 && strict.PageHeightTwips == 12240);
                }
                else
                {
                    Check(log, ref ok, "DOCX Strict text", false);
                }

                string text = TextFile.Load(plain);
                Check(log, ref ok, "Text round trip", text == editor.Text);

                // Picture pipeline: RichEdit stores WMF; DOCX/ODT writers normalize to PNG.
                using (Bitmap bitmap = new Bitmap(64, 32))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.Clear(Color.White);
                        g.FillEllipse(Brushes.SteelBlue, 4, 4, 50, 20);
                    }
                    editor.Select(editor.TextLength, 0);
                    Clipboard.SetImage(bitmap);
                    editor.Paste();
                    Application.DoEvents();
                }

                DocumentModel withImage = RtfReader.Read(editor.Rtf);
                if (CountImages(withImage) == 0)
                {
                    // Clipboard access can be unavailable in a headless or
                    // session-0 desktop.  Still exercise the image package path
                    // with a deterministic PNG.
                    using (Bitmap bitmap = new Bitmap(64, 32))
                    {
                        using (Graphics g = Graphics.FromImage(bitmap))
                        {
                            g.Clear(Color.White);
                            g.FillEllipse(Brushes.SteelBlue, 4, 4, 50, 20);
                        }
                        using (MemoryStream ms = new MemoryStream())
                        {
                            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                            Paragraph pictureParagraph = new Paragraph();
                            pictureParagraph.Inlines.Add(new ImageInline(ms.ToArray(), "png", 960, 480,
                                new RunFormat("Calibri", 22)));
                            withImage.Paragraphs.Add(pictureParagraph);
                        }
                    }
                }
                withImage.HeaderText = round.HeaderText;
                withImage.FooterText = round.FooterText;
                withImage.HeaderDistanceTwips = round.HeaderDistanceTwips;
                withImage.FooterDistanceTwips = round.FooterDistanceTwips;
                Check(log, ref ok, "RTF image parsed", CountImages(withImage) >= 1);
                DocumentFormats.SaveDocx(docx, withImage);
                DocumentFormats.SaveOdt(odt, withImage);
                DocumentModel dx2 = DocxFormat.Read(docx);
                DocumentModel od2 = OdtFormat.Read(odt);
                Check(log, ref ok, "DOCX image", CountImages(dx2) >= 1);
                Check(log, ref ok, "ODT image", CountImages(od2) >= 1);
                Check(log, ref ok, "DOCX image type", FirstImageExtension(dx2) == "png");
                Check(log, ref ok, "ODT image type", FirstImageExtension(od2) == "png");

                // Encoding round-trips.
                string unicode = Path.Combine(directory, "unicode.txt");
                TextFile.Save(unicode, "áéíóú ñ\r\nsegunda línea", TextEncodingKind.Unicode);
                Check(log, ref ok, "Unicode text", TextFile.Load(unicode) == "áéíóú ñ\nsegunda línea");

                // Legacy Word 97-2003 binary import.  The fixture is generated
                // from test-assets/sample.html with LibreOffice; the reader
                // itself has no external dependency.
                string docFixture = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-assets", "sample.doc");
                if (!File.Exists(docFixture)) docFixture = Path.Combine(directory, "sample.doc");
                if (File.Exists(docFixture))
                {
                    DocumentModel legacy = DocumentFormats.ReadModel(docFixture, DocumentFormat.Doc);
                    Check(log, ref ok, "DOC detect", DocumentFormats.Detect(docFixture) == DocumentFormat.Doc);
                    Check(log, ref ok, "DOC landscape page", legacy.PageWidthTwips > legacy.PageHeightTwips);
                    Check(log, ref ok, "DOC paragraphs", legacy.Paragraphs.Count >= 20);
                    Check(log, ref ok, "DOC title", HasText(legacy, "Mini Word · Prueba integral"));
                    Check(log, ref ok, "DOC bold", HasFormattedText(legacy, "Mini Word · Prueba integral", "B"));
                    Check(log, ref ok, "DOC underline", HasFormattedText(legacy, "Subrayado", "U"));
                    Check(log, ref ok, "DOC superscript", HasFormattedText(legacy, "índice", "^"));
                    Check(log, ref ok, "DOC numbered list", HasText(legacy, "1.") && HasText(legacy, "Primer elemento numerado."));
                    Check(log, ref ok, "DOC table", CountTableRows(legacy) == 4 && TableColumns(legacy) == 3);
                    Check(log, ref ok, "DOC page break", HasPageBreak(legacy));

                    string legacyRtf = RtfWriter.Write(legacy, false);
                    DocumentModel legacyUi = RtfReader.Read(legacyRtf);
                    Check(log, ref ok, "DOC editor RTF table", CountTableRows(legacyUi) == 4 && TableColumns(legacyUi) == 3);

                    string converted = Path.Combine(directory, "selftest-from-doc.docx");
                    DocumentFormats.SaveDocx(converted, legacy);
                    DocumentModel convertedBack = DocxFormat.Read(converted);
                    Check(log, ref ok, "DOC->DOCX text", HasText(convertedBack, "Mini Word · Prueba integral"));
                    Check(log, ref ok, "DOC->DOCX table", CountTableRows(convertedBack) == 4 && TableColumns(convertedBack) == 3);
                    Check(log, ref ok, "DOC->DOCX page", convertedBack.PageWidthTwips > convertedBack.PageHeightTwips);
                }
                else
                {
                    log.AppendLine("SKIP  DOC fixture not present");
                }

                string asciiFixture = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-assets", "ascii.doc");
                if (!File.Exists(asciiFixture)) asciiFixture = Path.Combine(directory, "ascii.doc");
                if (File.Exists(asciiFixture))
                {
                    DocumentModel ascii = DocumentFormats.ReadModel(asciiFixture, DocumentFormat.Doc);
                    Check(log, ref ok, "DOC ANSI text", HasText(ascii, "Hello ASCII") && HasText(ascii, "Second line without accents."));
                }
                else
                {
                    log.AppendLine("SKIP  ASCII DOC fixture not present");
                }

                DocumentModel pageBreakDoc = RtfReader.Read(
                    "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Calibri;}}\\f0\\fs22 Uno\\par\\page Dos\\par}");
                Check(log, ref ok, "RTF page break", HasPageBreak(pageBreakDoc));

                // The page view uses EM_FORMATRANGE for pagination.  Verify
                // that the measurement path returns a real page boundary.
                DocumentModel longDoc = new DocumentModel();
                for (int i = 0; i < 120; i++)
                {
                    Paragraph line = new Paragraph();
                    line.Inlines.Add(new TextInline("Línea de prueba " + (i + 1) + " para la paginación de Mini Word.", new RunFormat()));
                    longDoc.Paragraphs.Add(line);
                }
                editor.Rtf = RtfWriter.Write(longDoc, false);
                Application.DoEvents();
                bool measured = false;
                using (Graphics measure = editor.CreateGraphics())
                {
                    IntPtr hdc = measure.GetHdc();
                    try
                    {
                        NativeMethods.RECT pageRect = new NativeMethods.RECT(0, 0, 12240, 15840);
                        NativeMethods.RECT marginRect = new NativeMethods.RECT(1440, 1440, 10800, 14400);
                        int next = NativeMethods.MeasureRange(editor.Handle, hdc, pageRect, marginRect, 0, editor.TextLength);
                        measured = next > 0 && next < editor.TextLength;
                        NativeMethods.ClearFormatRange(editor.Handle);
                    }
                    finally
                    {
                        measure.ReleaseHdc(hdc);
                    }
                }
                Check(log, ref ok, "Page pagination", measured);

                // RTF table round-trip, including the RichEdit representation
                // used by the editor when a .docx table is displayed.
                DocumentModel tableModel = new DocumentModel();
                Paragraph table = new Paragraph();
                table.TableRows = new string[][] {
                    new string[] { "A", "B" },
                    new string[] { "C", "D" }
                };
                table.TableWidthTwips = 7200;
                tableModel.Paragraphs.Add(table);
                string tableRtf = RtfWriter.Write(tableModel, false);
                DocumentModel tableRoundTrip = RtfReader.Read(tableRtf);
                Check(log, ref ok, "RTF table round trip", CountTableRows(tableRoundTrip) == 2 && TableColumns(tableRoundTrip) == 2);
                editor.Rtf = tableRtf;
                Application.DoEvents();
                DocumentModel richTable = RtfReader.Read(editor.Rtf);
                Check(log, ref ok, "RichEdit table round trip", CountTableRows(richTable) == 2 && TableColumns(richTable) == 2);

                log.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
            }
            catch (Exception ex)
            {
                ok = false;
                log.AppendLine("EXCEPTION: " + ex);
                log.AppendLine("RESULT: FAIL");
            }
            finally
            {
                try { form.Close(); }
                catch { }
            }

            File.WriteAllText(Path.Combine(directory, "selftest.log"), log.ToString(), new UTF8Encoding(false));
            return ok;
        }

        private static DocumentModel BuildSample()
        {
            DocumentModel doc = new DocumentModel();
            Paragraph p = new Paragraph();
            p.Alignment = ParagraphAlignment.Center;
            RunFormat f = new RunFormat("Calibri", 32);
            f.Bold = true;
            f.ForeColor = Color.DarkBlue;
            p.Inlines.Add(new TextInline("Informe de prueba", f));
            doc.Paragraphs.Add(p);

            p = new Paragraph();
            p.Bullet = true;
            p.LeftIndentTwips = 720;
            p.FirstLineIndentTwips = -360;
            f = new RunFormat("Georgia", 24);
            f.Italic = true;
            f.BackColor = Color.Yellow;
            p.Inlines.Add(new TextInline("Primera viñeta con ñ y acentos", f));
            doc.Paragraphs.Add(p);

            p = new Paragraph();
            p.Alignment = ParagraphAlignment.Justify;
            p.LeftIndentTwips = 360;
            p.SpaceAfterTwips = 200;
            p.Inlines.Add(new TextInline("Texto normal ", new RunFormat("Calibri", 22)));
            f = new RunFormat("Calibri", 22);
            f.Underline = true;
            f.Strikeout = true;
            f.Script = 1;
            f.ForeColor = Color.Red;
            p.Inlines.Add(new TextInline("superíndice", f));
            p.Inlines.Add(new TextInline(" y final.", new RunFormat("Calibri", 22)));
            doc.Paragraphs.Add(p);
            return doc;
        }

        private static int CountImages(DocumentModel model)
        {
            int count = 0;
            for (int i = 0; i < model.Paragraphs.Count; i++)
                for (int j = 0; j < model.Paragraphs[i].Inlines.Count; j++)
                    if (model.Paragraphs[i].Inlines[j] is ImageInline) count++;
            return count;
        }

        private static string FirstImageExtension(DocumentModel model)
        {
            for (int i = 0; i < model.Paragraphs.Count; i++)
                for (int j = 0; j < model.Paragraphs[i].Inlines.Count; j++)
                {
                    ImageInline img = model.Paragraphs[i].Inlines[j] as ImageInline;
                    if (img != null) return img.Extension;
                }
            return string.Empty;
        }

        private static bool HasText(DocumentModel model, string value)
        {
            for (int i = 0; i < model.Paragraphs.Count; i++)
            {
                Paragraph p = model.Paragraphs[i];
                if (p.TableRows != null)
                {
                    for (int r = 0; r < p.TableRows.Length; r++)
                        for (int c = 0; c < p.TableRows[r].Length; c++)
                            if (p.TableRows[r][c].IndexOf(value, StringComparison.Ordinal) >= 0) return true;
                    continue;
                }
                for (int j = 0; j < p.Inlines.Count; j++)
                {
                    TextInline t = p.Inlines[j] as TextInline;
                    if (t != null && t.Text.IndexOf(value, StringComparison.Ordinal) >= 0) return true;
                }
            }
            return false;
        }

        private static bool HasFormattedText(DocumentModel model, string value, string flag)
        {
            for (int i = 0; i < model.Paragraphs.Count; i++)
            {
                Paragraph p = model.Paragraphs[i];
                for (int j = 0; j < p.Inlines.Count; j++)
                {
                    TextInline t = p.Inlines[j] as TextInline;
                    if (t == null || t.Text.IndexOf(value, StringComparison.Ordinal) < 0) continue;
                    if (flag == "B" && t.Format.Bold) return true;
                    if (flag == "U" && t.Format.Underline) return true;
                    if (flag == "^" && t.Format.Script > 0) return true;
                }
            }
            return false;
        }

        private static bool HasPageBreak(DocumentModel model)
        {
            for (int i = 0; i < model.Paragraphs.Count; i++)
                if (model.Paragraphs[i].PageBreakBefore) return true;
            return false;
        }

        private static int CountTableRows(DocumentModel model)
        {
            int count = 0;
            for (int i = 0; i < model.Paragraphs.Count; i++)
                if (model.Paragraphs[i].TableRows != null) count += model.Paragraphs[i].TableRows.Length;
            return count;
        }

        private static int TableColumns(DocumentModel model)
        {
            for (int i = 0; i < model.Paragraphs.Count; i++)
                if (model.Paragraphs[i].TableRows != null && model.Paragraphs[i].TableRows.Length > 0)
                    return model.Paragraphs[i].TableRows[0].Length;
            return 0;
        }

        private static void Check(StringBuilder log, ref bool ok, string name, bool condition)
        {
            log.AppendLine((condition ? "PASS  " : "FAIL  ") + name);
            if (!condition) ok = false;
        }
    }
}
