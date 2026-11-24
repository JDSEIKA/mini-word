using System;
using System.Collections.Generic;
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

        DocumentModel advanced = new DocumentModel();
        Paragraph advancedPara = new Paragraph();
        TextInline pageField = new TextInline("1", new RunFormat("Calibri", 22));
        pageField.Field = FieldKind.Page;
        pageField.FieldInstruction = " PAGE ";
        advancedPara.Inlines.Add(pageField);
        TextInline inserted = new TextInline("nuevo", new RunFormat("Calibri", 22));
        inserted.Revision = RevisionKind.Insert;
        inserted.RevisionAuthor = "Tester";
        advancedPara.Inlines.Add(inserted);
        TextInline deleted = new TextInline("viejo", new RunFormat("Calibri", 22));
        deleted.Revision = RevisionKind.Delete;
        advancedPara.Inlines.Add(deleted);
        advanced.Paragraphs.Add(advancedPara);
        string advancedPath = "tmp-regression/advanced-audit.docx";
        if (File.Exists(advancedPath)) File.Delete(advancedPath);
        DocumentFormats.SaveDocx(advancedPath, advanced);
        DocumentModel advancedRound = DocxFormat.Read(advancedPath);
        bool fieldRound = false, insertRound = false, deleteRound = false;
        for (int i = 0; i < advancedRound.Paragraphs[0].Inlines.Count; i++)
        {
            TextInline run = advancedRound.Paragraphs[0].Inlines[i] as TextInline;
            if (run == null) continue;
            if (run.Field == FieldKind.Page) fieldRound = true;
            if (run.Revision == RevisionKind.Insert) insertRound = true;
            if (run.Revision == RevisionKind.Delete) deleteRound = true;
        }
        if (!fieldRound || !insertRound || !deleteRound) throw new Exception("field/revision roundtrip");

        DocumentModel oldCompare = new DocumentModel();
        Paragraph oldLine = new Paragraph();
        oldLine.OutlineLevel = 1;
        oldLine.Inlines.Add(new TextInline("Hola mundo viejo", new RunFormat("Calibri", 22)));
        oldCompare.Paragraphs.Add(oldLine);
        DocumentModel newCompare = new DocumentModel();
        Paragraph newLine = new Paragraph();
        newLine.OutlineLevel = 1;
        newLine.Inlines.Add(new TextInline("Hola mundo nuevo", new RunFormat("Calibri", 22)));
        newCompare.Paragraphs.Add(newLine);
        DocumentModel compared = DocumentCompare.Compare(oldCompare, newCompare);
        bool comparedInsert = false, comparedDelete = false;
        for (int i = 0; i < compared.Paragraphs[0].Inlines.Count; i++)
        {
            TextInline run = compared.Paragraphs[0].Inlines[i] as TextInline;
            if (run == null) continue;
            if (run.Revision == RevisionKind.Insert) comparedInsert = true;
            if (run.Revision == RevisionKind.Delete) comparedDelete = true;
        }
        if (!comparedInsert || !comparedDelete) throw new Exception("document compare revisions");
        List<Paragraph> toc = TableOfContents.Build(newCompare, 3, delegate(int index) { return 1; });
        if (toc.Count < 2 || !toc[1].IsTocEntry) throw new Exception("table of contents");
        DocumentModel tableRuns = new DocumentModel();
        Paragraph tableRunsPara = new Paragraph();
        tableRunsPara.TableRows = new string[][] { new string[] { "ab", "cd" } };
        tableRunsPara.TableWidthTwips = 7000;
        tableRunsPara.RepeatTableHeader = true;
        tableRunsPara.EnsureTableCells();
        TableCell runCell = tableRunsPara.TableCells[0][0];
        runCell.Runs.Clear();
        RunFormat boldRun = new RunFormat("Calibri", 22); boldRun.Bold = true;
        TableCellRun firstCellRun = new TableCellRun("a", boldRun);
        TableCellRun secondCellRun = new TableCellRun("b", new RunFormat("Calibri", 22));
        secondCellRun.LinkUrl = "https://example.com";
        runCell.Runs.Add(firstCellRun); runCell.Runs.Add(secondCellRun); runCell.SyncFromRuns();
        tableRuns.Paragraphs.Add(tableRunsPara);
        tableRuns.Comments.Add(new DocumentComment
        {
            Id = 7, Author = "Tester", Text = "Comentario",
            Replies = { new CommentReply { Text = "Respuesta", Author = "Tester", Date = "2026-11-23" } }
        });
        string tableRunsPath = "tmp-regression/table-runs-audit.docx";
        if (File.Exists(tableRunsPath)) File.Delete(tableRunsPath);
        DocumentFormats.SaveDocx(tableRunsPath, tableRuns);
        DocumentModel tableRunsRound = DocxFormat.Read(tableRunsPath);
        TableCell roundCell = tableRunsRound.Paragraphs[0].TableCells[0][0];
        if (roundCell.Runs.Count != 2 || string.IsNullOrEmpty(roundCell.Runs[1].LinkUrl))
            throw new Exception("table cell runs roundtrip");
        if (!tableRunsRound.Paragraphs[0].RepeatTableHeader)
            throw new Exception("repeat table header roundtrip");
        if (tableRunsRound.Comments.Count != 1 || tableRunsRound.Comments[0].Replies.Count != 1)
            throw new Exception("comment replies roundtrip");
        editor.ApplyParagraphStyle("Título 1");
        if (editor.Document.Paragraphs[0].StyleName != "Título 1")
            throw new Exception("named paragraph style");
        form.Close();
        Console.WriteLine("FEATURE_AUDIT=PASS");
    }
}
