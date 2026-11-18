using System;
using System.Drawing;
using System.IO;
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
        paragraph.Inlines.Add(new TextInline("url", new RunFormat("Arial", 22)));
        doc.Paragraphs.Add(paragraph);
        form.Show();
        editor.Document = doc;
        Application.DoEvents();

        editor.SelectRange(0, 0, 3);
        if (!editor.ApplyHyperlink("https://example.com"))
            throw new Exception("apply hyperlink");
        if (((TextInline)doc.Paragraphs[0].Inlines[0]).Format.LinkUrl != "https://example.com")
            throw new Exception("hyperlink formatting");

        Paragraph table = new Paragraph();
        table.TableRows = new string[][] { new string[] { "A", "B" }, new string[] { "C", "D" } };
        table.EnsureTableCells();
        table.TableCells[0][0].LinkUrl = "https://example.org";
        doc.Paragraphs.Add(table);

        FloatingObject first = new FloatingObject
        {
            Kind = FloatingKind.Shape,
            ShapeKind = FloatingShapeKind.Rectangle
        };
        FloatingObject second = new FloatingObject
        {
            Kind = FloatingKind.Shape,
            ShapeKind = FloatingShapeKind.Ellipse
        };
        doc.Paragraphs[0].FloatingObjects.Add(first);
        doc.Paragraphs[0].FloatingObjects.Add(second);
        typeof(NativeDocumentEditor).GetField("_selectedFloating",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
            .SetValue(editor, first);
        if (!editor.BringSelectedFloatingToFront()
            || doc.Paragraphs[0].FloatingObjects[1] != first)
            throw new Exception("bring to front");
        if (!editor.SendSelectedFloatingToBack()
            || doc.Paragraphs[0].FloatingObjects[0] != first)
            throw new Exception("send to back");
        if (!editor.RotateSelectedFloating(90) || first.Rotation != 90)
            throw new Exception("rotate object");

        string path = "tmp-regression/feature-audit.docx";
        if (File.Exists(path)) File.Delete(path);
        if (File.Exists(path + ".bak")) File.Delete(path + ".bak");
        DocumentFormats.SaveDocx(path, doc);
        DocumentModel reopened = DocxFormat.Read(path);
        bool tableLink = false;
        for (int i = 0; i < reopened.Paragraphs.Count; i++)
        {
            TableCell[][] cells = reopened.Paragraphs[i].TableCells;
            if (cells == null) continue;
            for (int r = 0; r < cells.Length; r++)
                for (int c = 0; c < cells[r].Length; c++)
                    if (cells[r][c] != null && cells[r][c].LinkUrl == "https://example.org")
                        tableLink = true;
        }
        if (!tableLink) throw new Exception("table hyperlink roundtrip");
        bool rotationRoundtrip = false;
        for (int i = 0; i < reopened.Paragraphs.Count; i++)
            for (int j = 0; j < reopened.Paragraphs[i].FloatingObjects.Count; j++)
                if (reopened.Paragraphs[i].FloatingObjects[j].Rotation == 90)
                    rotationRoundtrip = true;
        if (!rotationRoundtrip) throw new Exception("rotation roundtrip");

        form.Close();
        Console.WriteLine("FEATURE_AUDIT=PASS");
    }
}
