using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using WordPadClone;

internal static class P
{
    [STAThread]
    private static void Main()
    {
        string root = Path.Combine("tmp-regression", "autosave-audit");
        Directory.CreateDirectory(root);
        string original = Path.Combine(root, "original.docx");
        File.WriteAllBytes(original, new byte[] { 1, 2, 3, 4, 5 });
        string recoveryRoot = Path.Combine(root, "recovery");
        Environment.SetEnvironmentVariable("MINIWORD_RECOVERY_DIR", recoveryRoot);
        NativeEditorForm form = new NativeEditorForm(new string[0]);
        try
        {
            Type type = typeof(NativeEditorForm);
            type.GetField("_path", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, original);
            type.GetField("_dirty", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, true);
            type.GetField("_autoSaveEnabled", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, true);
            type.GetMethod("AutoSaveNow", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, null);
            byte[] after = File.ReadAllBytes(original);
            if (after.Length != 5 || after[0] != 1) throw new Exception("autosave modified original");
            if (Directory.GetFiles(recoveryRoot, "recovery-*.docx").Length == 0)
                throw new Exception("autosave recovery missing");
            type.GetField("_autoSaveEnabled", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, false);
            type.GetField("_dirty", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, true);
            type.GetMethod("AutoSaveNow", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, null);
            if (Directory.GetFiles(recoveryRoot, "recovery-*.docx").Length == 0)
                throw new Exception("recovery missing when autosave disabled");
            Console.WriteLine("AUTOSAVE_AUDIT=PASS");
        }
        finally
        {
            form.Dispose();
        }
    }
}
