using System;
using System.Collections;
using System.Drawing;
using System.Reflection;
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
        paragraph.Inlines.Add(new TextInline("hola mundo", new RunFormat("Calibri", 22)));
        doc.Paragraphs.Add(paragraph);

        form.Show();
        editor.Document = doc;
        Application.DoEvents();

        Type type = typeof(NativeDocumentEditor);
        type.GetMethod("EnsureLayout", BindingFlags.Instance | BindingFlags.NonPublic)
            .Invoke(editor, null);
        IList lines = (IList)type.GetField("_lines", BindingFlags.Instance | BindingFlags.NonPublic)
            .GetValue(editor);
        object line = lines[0];
        IList segments = (IList)line.GetType().GetField("Segments").GetValue(line);
        object segment = null;
        for (int i = 0; i < segments.Count; i++)
        {
            object candidate = segments[i];
            string text = (string)candidate.GetType().GetField("Text").GetValue(candidate);
            if (text == "hola") { segment = candidate; break; }
        }
        if (segment == null) throw new Exception("caret: segment not found");

        float x = (float)segment.GetType().GetField("X").GetValue(segment);
        float width = (float)segment.GetType().GetField("Width").GetValue(segment);
        float y = (float)segment.GetType().GetField("Y").GetValue(segment);
        float height = (float)segment.GetType().GetField("Height").GetValue(segment);
        Point click = new Point((int)(x + width * 0.75f), (int)(y + height / 2f));

        MethodInfo pointToPosition = type.GetMethod("PointToPosition",
            BindingFlags.Instance | BindingFlags.NonPublic);
        object caret = pointToPosition.Invoke(editor, new object[] { click });
        int offset = (int)caret.GetType().GetField("Offset").GetValue(caret);
        if (offset != 3) throw new Exception("caret: intra-word offset was " + offset);

        bool rendered = false;
        foreach (MethodInfo candidate in type.GetMethods(BindingFlags.Instance | BindingFlags.NonPublic))
        {
            if (candidate.Name != "XForOffset" || candidate.GetParameters().Length != 2
                || candidate.GetParameters()[0].ParameterType.Name != "LayoutSegment") continue;
            float caretX = (float)candidate.Invoke(editor, new object[] { segment, 3 });
            if (caretX <= x + width * 0.25f || caretX >= x + width)
                throw new Exception("caret: rendered X was " + caretX);
            rendered = true;
            break;
        }
        if (!rendered) throw new Exception("caret: XForOffset not found");
        MethodInfo mouseDown = type.GetMethod("OnMouseDown",
            BindingFlags.Instance | BindingFlags.NonPublic);
        mouseDown.Invoke(editor, new object[] { new MouseEventArgs(MouseButtons.Left, 1, click.X, click.Y, 0) });
        object eventCaret = type.GetField("_caret", BindingFlags.Instance | BindingFlags.NonPublic)
            .GetValue(editor);
        int eventOffset = (int)eventCaret.GetType().GetField("Offset").GetValue(eventCaret);
        if (eventOffset != 3) throw new Exception("caret: mouse event offset was " + eventOffset);

        form.Close();
        Console.WriteLine("CARET_SMOKE=PASS");
    }
}
