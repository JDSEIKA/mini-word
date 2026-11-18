using System;
using System.Drawing;
using System.Windows.Forms;
using WordPadClone.Controls;
using WordPadClone.Formats;

internal static class P
{
    [STAThread]
    private static void Main()
    {
        Form form = new Form();
        form.Size = new Size(900, 700);
        NativeDocumentEditor editor = new NativeDocumentEditor();
        editor.Dock = DockStyle.Fill;
        form.Controls.Add(editor);
        DocumentModel doc = new DocumentModel();
        Paragraph paragraph = new Paragraph();
        paragraph.Inlines.Add(new TextInline("Normal ", new RunFormat("Arial", 22)));
        paragraph.Inlines.Add(new TextInline("Bold", new RunFormat("Arial", 32) { Bold = true }));
        paragraph.Inlines.Add(new TextInline(" Under", new RunFormat("Arial", 22) { Underline = true }));
        doc.Paragraphs.Add(paragraph);
        form.Show();
        editor.Document = doc;
        Application.DoEvents();

        editor.SelectAll();
        editor.Copy();
        DocumentModel target = new DocumentModel();
        Paragraph pasted = new Paragraph();
        target.Paragraphs.Add(pasted);
        editor.Document = target;
        editor.PlaceCaret(0, 0);
        editor.Paste();
        if (!HasRun(pasted, "Bold", 32, true, false)
            || !HasRun(pasted, " Under", 22, false, true))
            throw new Exception("internal rich clipboard");

        string html = "<html><body><!--StartFragment--><span style=\"font-size:24px;font-weight:bold;text-decoration:underline\">HTML BIG</span> normal<!--EndFragment--></body></html>";
        DataObject data = new DataObject();
        data.SetData(DataFormats.Html, html);
        Clipboard.SetDataObject(data, true);
        Application.DoEvents();
        System.Threading.Thread.Sleep(100);
        Paragraph htmlParagraph = new Paragraph();
        target.Paragraphs.Add(htmlParagraph);
        editor.PlaceCaret(1, 0);
        editor.Paste();
        if (!HasRun(htmlParagraph, "HTML BIG", 36, true, true))
            throw new Exception("HTML rich clipboard");

        Clipboard.SetText("PLAIN");
        Paragraph plainParagraph = new Paragraph();
        target.Paragraphs.Add(plainParagraph);
        editor.SetFontFamily("Arial");
        editor.SetFontSize(30);
        editor.PlaceCaret(2, 0);
        editor.Paste();
        if (!HasRun(plainParagraph, "PLAIN", 22, false, false)
            || ((TextInline)plainParagraph.Inlines[0]).Format.FontFamily != "Calibri")
            throw new Exception("plain neutral clipboard");

        string broken = "<html><body><!--StartFragment--><p>tambiÃÂ©n, quÃÂ©, estÃÂ¡n y ediciÃÂÃ³n</p><!--EndFragment--></body></html>";
        DataObject brokenData = new DataObject();
        brokenData.SetData(DataFormats.Html, broken);
        Clipboard.SetDataObject(brokenData, true);
        Application.DoEvents();
        System.Threading.Thread.Sleep(100);
        Paragraph fixedParagraph = new Paragraph();
        target.Paragraphs.Add(fixedParagraph);
        editor.PlaceCaret(3, 0);
        editor.Paste();
        string fixedText = ((TextInline)fixedParagraph.Inlines[0]).Text;
        if (!fixedText.Contains("también") || !fixedText.Contains("qué")
            || !fixedText.Contains("están") || !fixedText.Contains("edición"))
            throw new Exception("mojibake clipboard: " + fixedText);

        Clipboard.SetText("SIN FORMATO");
        Application.DoEvents();
        System.Threading.Thread.Sleep(100);
        Paragraph plainOnlyParagraph = new Paragraph();
        target.Paragraphs.Add(plainOnlyParagraph);
        editor.SetFontFamily("Arial");
        editor.SetFontSize(30);
        editor.PlaceCaret(4, 0);
        editor.PastePlainText();
        if (!HasRun(plainOnlyParagraph, "SIN FORMATO", 22, false, false)
            || ((TextInline)plainOnlyParagraph.Inlines[0]).Format.FontFamily != "Calibri")
            throw new Exception("paste without formatting");

        using (Bitmap clipboardImage = new Bitmap(24, 12))
        {
            Clipboard.SetImage(clipboardImage);
            Application.DoEvents();
            System.Threading.Thread.Sleep(100);
            Paragraph imageParagraph = new Paragraph();
            target.Paragraphs.Add(imageParagraph);
            editor.PlaceCaret(5, 0);
            editor.Paste();
            if (imageParagraph.FloatingObjects.Count != 1
                || imageParagraph.FloatingObjects[0].Kind != FloatingKind.Image
                || imageParagraph.FloatingObjects[0].Image == null
                || imageParagraph.FloatingObjects[0].Image.Data.Length == 0)
                throw new Exception("clipboard image not embedded");
        }

        form.Close();
        Console.WriteLine("CLIPBOARD_SMOKE=PASS");
    }

    private static bool HasRun(Paragraph paragraph, string text, int size, bool bold, bool underline)
    {
        for (int i = 0; i < paragraph.Inlines.Count; i++)
        {
            TextInline run = paragraph.Inlines[i] as TextInline;
            if (run == null || run.Text.IndexOf(text, StringComparison.Ordinal) < 0) continue;
            return run.Format.SizeHalfPoints == size && run.Format.Bold == bold
                && run.Format.Underline == underline;
        }
        return false;
    }
}
