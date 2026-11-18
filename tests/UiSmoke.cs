using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using WordPadClone.Controls;
using WordPadClone.Dialogs;
using WordPadClone.Formats;

internal static class P
{
    [STAThread]
    private static void Main()
    {
        HeaderFooterDialog dialog = new HeaderFooterDialog(
            "Normal H", "Normal F", "First H", "First F", "Even H", "Even F",
            720, 720, new RunFormat("Arial", 22), ParagraphAlignment.Left, true, true);
        if (!dialog.DifferentFirstPage || !dialog.DifferentOddEven
            || dialog.FirstPageHeaderText != "First H" || dialog.EvenPageFooterText != "Even F")
            throw new Exception("header dialog");

        byte[] pixel = File.ReadAllBytes("test-assets/pixel.png");
        List<ImageInline> firstImages = new List<ImageInline>();
        firstImages.Add(new ImageInline(pixel, "png", 300, 300, new RunFormat("Arial", 18)));
        HeaderFooterDialog variants = new HeaderFooterDialog(
            "N", "NF", "F", "FF", "E", "EF", 720, 720,
            new RunFormat[] { new RunFormat("Arial", 22), new RunFormat("Verdana", 72), new RunFormat("Courier New", 60) },
            new RunFormat[] { new RunFormat("Arial", 22), new RunFormat("Verdana", 72), new RunFormat("Courier New", 60) },
            new ParagraphAlignment[] { ParagraphAlignment.Left, ParagraphAlignment.Center, ParagraphAlignment.Right },
            new ParagraphAlignment[] { ParagraphAlignment.Left, ParagraphAlignment.Center, ParagraphAlignment.Right },
            new List<ImageInline>[] { new List<ImageInline>(), firstImages, new List<ImageInline>() },
            new List<ImageInline>[] { new List<ImageInline>(), new List<ImageInline>(), new List<ImageInline>() },
            true, true);
        if (variants.GetHeaderFormat(1).SizeHalfPoints != 72
            || variants.GetHeaderFormat(2).SizeHalfPoints != 60
            || variants.GetHeaderAlignment(2) != ParagraphAlignment.Right
            || variants.GetHeaderImages(1).Count != 1)
            throw new Exception("variant format dialog");
        variants.Dispose();

        NativeDocumentEditor editor = new NativeDocumentEditor();
        Form form = new Form();
        form.Size = new Size(900, 700);
        editor.Dock = DockStyle.Fill;
        form.Controls.Add(editor);
        form.Show();
        try
        {
            DocumentModel doc = new DocumentModel();
            for (int i = 0; i < 3; i++)
            {
                Paragraph p = new Paragraph();
                p.Inlines.Add(new TextInline("Párrafo " + i, new RunFormat("Arial", 22)));
                doc.Paragraphs.Add(p);
            }
            editor.Document = doc;
            editor.SelectRange(1, 0, 3);
            editor.SetFontSize(30);
            if (((TextInline)doc.Paragraphs[0].Inlines[0]).Format.SizeHalfPoints != 22
                || ((TextInline)doc.Paragraphs[2].Inlines[0]).Format.SizeHalfPoints != 22
                || ((TextInline)doc.Paragraphs[1].Inlines[0]).Format.SizeHalfPoints != 60)
                throw new Exception("format scope");

            editor.InsertShape(FloatingShapeKind.Triangle);
            int insertedShapeParagraph = -1;
            for (int i = 0; i < doc.Paragraphs.Count; i++)
                if (doc.Paragraphs[i].FloatingObjects.Count > 0) insertedShapeParagraph = i;
            if (insertedShapeParagraph != 1) throw new Exception("shape insertion page");
            editor.Copy();
            editor.Paste();
            int shapeCount = 0;
            for (int i = 0; i < doc.Paragraphs.Count; i++)
                for (int j = 0; j < doc.Paragraphs[i].FloatingObjects.Count; j++)
                    if (doc.Paragraphs[i].FloatingObjects[j].Kind == FloatingKind.Shape) shapeCount++;
            if (shapeCount != 2) throw new Exception("shape paste");
            editor.Cut();
            shapeCount = 0;
            for (int i = 0; i < doc.Paragraphs.Count; i++)
                for (int j = 0; j < doc.Paragraphs[i].FloatingObjects.Count; j++)
                    if (doc.Paragraphs[i].FloatingObjects[j].Kind == FloatingKind.Shape) shapeCount++;
            if (shapeCount != 1) throw new Exception("shape cut");

            FloatingObject textBox = new FloatingObject
            {
                Kind = FloatingKind.TextBox,
                Text = "Color",
                Format = new RunFormat("Arial", 22)
            };
            doc.Paragraphs[0].FloatingObjects.Add(textBox);
            typeof(NativeDocumentEditor).GetField("_selectedFloating",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .SetValue(editor, textBox);
            editor.SetForeColor(Color.Red);
            if (!textBox.Format.ForeColor.HasValue || textBox.Format.ForeColor.Value != Color.Red)
                throw new Exception("textbox color");
        }
        finally
        {
            form.Close();
            dialog.Dispose();
        }

        DocumentModel shapeDoc = new DocumentModel();
        Paragraph shapeParagraph = new Paragraph();
        shapeParagraph.FloatingObjects.Add(new FloatingObject
        {
            Kind = FloatingKind.Shape,
            ShapeKind = FloatingShapeKind.Triangle,
            WidthTwips = 2000,
            HeightTwips = 1600
        });
        shapeDoc.Paragraphs.Add(shapeParagraph);
        string path = "tmp-regression/ui-shape.docx";
        if (File.Exists(path)) File.Delete(path);
        if (File.Exists(path + ".bak")) File.Delete(path + ".bak");
        DocumentFormats.SaveDocx(path, shapeDoc);
        DocumentModel reopened = DocxFormat.Read(path);
        if (reopened.Paragraphs[0].FloatingObjects.Count != 1
            || reopened.Paragraphs[0].FloatingObjects[0].ShapeKind != FloatingShapeKind.Triangle)
            throw new Exception("shape roundtrip");

        Console.WriteLine("UI_SMOKE=PASS");
    }
}
