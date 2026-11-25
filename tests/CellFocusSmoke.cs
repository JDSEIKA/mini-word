using System;
using System.Collections;
using System.Reflection;
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
            Type formType = typeof(NativeEditorForm);
            NativeDocumentEditor editor = (NativeDocumentEditor)formType.GetField("_editor",
                BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
            DocumentModel document = new DocumentModel();
            Paragraph paragraph = new Paragraph();
            paragraph.TableRows = new string[][] { new string[] { "abc def" } };
            paragraph.TableWidthTwips = 3000;
            paragraph.EnsureTableCells();
            document.Paragraphs.Add(paragraph);
            editor.Document = document;
            Application.DoEvents();

            object cell = null;
            IEnumerable tables = (IEnumerable)typeof(NativeDocumentEditor).GetField("_tables",
                BindingFlags.NonPublic | BindingFlags.Instance).GetValue(editor);
            foreach (object table in tables)
            {
                IEnumerable cells = (IEnumerable)table.GetType().GetField("Cells").GetValue(table);
                foreach (object candidate in cells) { cell = candidate; break; }
                if (cell != null) break;
            }
            if (cell == null) throw new Exception("table cell not laid out");

            MethodInfo start = typeof(NativeDocumentEditor).GetMethod("StartCellEditor",
                BindingFlags.NonPublic | BindingFlags.Instance, Type.DefaultBinder,
                new Type[] { cell.GetType() }, null);
            start.Invoke(editor, new object[] { cell });

            RichTextBox box = (RichTextBox)typeof(NativeDocumentEditor).GetField("_cellEditor",
                BindingFlags.NonPublic | BindingFlags.Instance).GetValue(editor);
            box.SelectionStart = 0;
            box.SelectionLength = 3;
            ToolStripButton bold = (ToolStripButton)formType.GetField("_boldButton",
                BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
            bold.PerformClick();
            Application.DoEvents();
            if (!editor.IsEditingTableCell) throw new Exception("bold closed cell editor");
            if (box.SelectionStart != 0 || box.SelectionLength != 3) throw new Exception("bold lost cell selection");
            object cellEdit = typeof(NativeDocumentEditor).GetField("_cellEdit",
                BindingFlags.NonPublic | BindingFlags.Instance).GetValue(editor);
            object cellModel = cellEdit.GetType().GetField("Cell").GetValue(cellEdit);
            IList runs = (IList)cellModel.GetType().GetField("Runs").GetValue(cellModel);
            object firstFormat = runs[0].GetType().GetField("Format").GetValue(runs[0]);
            if (!(bool)firstFormat.GetType().GetField("Bold").GetValue(firstFormat))
                throw new Exception("bold not applied to cell run");

            ToolStripButton center = (ToolStripButton)formType.GetField("_centerButton",
                BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
            center.PerformClick();
            Application.DoEvents();
            if (!editor.IsEditingTableCell) throw new Exception("alignment closed cell editor");
            object alignment = cellModel.GetType().GetField("Alignment").GetValue(cellModel);
            if (alignment.ToString() != "Center") throw new Exception("cell alignment not applied");

            Console.WriteLine("CELL_FOCUS_SMOKE=PASS");
        }
        finally
        {
            form.Dispose();
        }
    }
}
