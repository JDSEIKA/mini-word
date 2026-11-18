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
