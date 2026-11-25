using System;
using System.Collections.Generic;
using System.Windows.Forms;
using WordPadClone;
using WordPadClone.Controls;
using WordPadClone.Formats;

internal static class P
{
    [STAThread]
    private static void Main()
    {
        NativeEditorForm form = new NativeEditorForm(new string[0]);
        form.Show();
        try
        {
            NativeDocumentEditor editor = (NativeDocumentEditor)typeof(NativeEditorForm).GetField("_editor",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(form);
            DocumentModel document = new DocumentModel();
            Paragraph paragraph = new Paragraph();
            paragraph.TableRows = new string[][] { new string[] { "hola mundo" } };
            paragraph.EnsureTableCells();
            document.Paragraphs.Add(paragraph);
            editor.Document = document;
            Application.DoEvents();

            string[] texts = editor.GetSpellParagraphs();
            int index = -1;
            for (int i = 0; i < texts.Length; i++) if (texts[i] == "hola mundo") { index = i; break; }
            if (index < 0) throw new Exception("table text not included in spell paragraphs");

            SpellError error = new SpellError();
            error.Paragraph = index;
            error.Start = 0;
            error.Length = 4;
            error.Word = "hola";
            List<SpellError> errors = new List<SpellError>();
            errors.Add(error);
            editor.ResolveSpellTargets(errors);
            if (error.TableRow != 0 || error.TableColumn != 0 || error.Paragraph != 0)
                throw new Exception("table spell target not resolved");

            if (!editor.ReplaceSpelling(error, "adiós")) throw new Exception("table spell replace failed");
            if (paragraph.TableCells[0][0].Text != "adiós mundo")
                throw new Exception("table spell text mismatch: " + paragraph.TableCells[0][0].Text);

            Console.WriteLine("SPELL_TABLE_SMOKE=PASS");
        }
        finally
        {
            form.Dispose();
        }
    }
}
