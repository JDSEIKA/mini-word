using System;
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
            Type type = typeof(NativeEditorForm);
            MethodInfo process = type.GetMethod("ProcessCmdKey", BindingFlags.NonPublic | BindingFlags.Instance);
            NativeDocumentEditor editor = (NativeDocumentEditor)type.GetField("_editor",
                BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
            Func<Keys, bool> send = delegate(Keys key)
            {
                Message message = new Message();
                object[] args = new object[] { message, key };
                return (bool)process.Invoke(form, args);
            };

            editor.Focus();
            editor.InsertText("abc def");
            if (!send(Keys.Control | Keys.E) || editor.CurrentAlignment != ParagraphAlignment.Center)
                throw new Exception("Ctrl+E");
            if (!send(Keys.Control | Keys.Shift | Keys.L) || !editor.CurrentParagraph.Bullet)
                throw new Exception("Ctrl+Shift+L");
            if (!send(Keys.Control | Keys.Shift | Keys.D7) || !editor.CurrentParagraph.Numbered)
                throw new Exception("Ctrl+Shift+7");
            if (!send(Keys.Control | Keys.D1) || editor.CurrentParagraph.LineSpacing != 0)
                throw new Exception("Ctrl+1");
            if (!send(Keys.Control | Keys.D5) || editor.CurrentParagraph.LineSpacing != 1)
                throw new Exception("Ctrl+5");
            if (!send(Keys.Control | Keys.D2) || editor.CurrentParagraph.LineSpacing != 2)
                throw new Exception("Ctrl+2");
            if (!send(Keys.Control | Keys.Alt | Keys.D1) || editor.CurrentParagraph.StyleName != "Título 1")
                throw new Exception("Ctrl+Alt+1");
            if (!send(Keys.Control | Keys.Shift | Keys.N) || editor.CurrentParagraph.StyleName != "Normal")
                throw new Exception("Ctrl+Shift+N");
            if (!send(Keys.Control | Keys.Shift | Keys.Space) || editor.Document.PlainText.IndexOf('\u00A0') < 0)
                throw new Exception("Ctrl+Shift+Space");
            int beforeSize = editor.CurrentFormat.SizeHalfPoints;
            if (!send(Keys.Control | Keys.Shift | Keys.OemPeriod) || editor.CurrentFormat.SizeHalfPoints <= beforeSize)
                throw new Exception("Ctrl+Shift+>");
            if (!send(Keys.Control | Keys.D0) || editor.CurrentParagraph.StyleName != "Normal")
                throw new Exception("Ctrl+0");

            string before = editor.Document.PlainText;
            editor.InsertLineBreak();
            if (editor.Document.Paragraphs.Count != 1 || editor.Document.PlainText.IndexOf('\n') < 0)
                throw new Exception("Shift+Enter line break");
            editor.DeleteWordBackward();
            if (editor.Document.PlainText == before) throw new Exception("Ctrl+Backspace");

            Console.WriteLine("SHORTCUT_SMOKE=PASS");
        }
        finally
        {
            form.Dispose();
        }
    }
}
