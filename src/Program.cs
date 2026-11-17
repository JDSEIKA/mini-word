using System;
using System.IO;
using System.Windows.Forms;
using WordPadClone.Formats;
using WordPadClone.Util;

namespace WordPadClone
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args != null && args.Length > 0 && string.Equals(args[0], "--selftest", StringComparison.OrdinalIgnoreCase))
            {
                string dir = args.Length > 1 ? args[1] : Path.Combine(Path.GetTempPath(), "MiniWordSelfTest");
                try
                {
                    bool ok = SelfTest.Run(dir);
                    Environment.Exit(ok ? 0 : 2);
                    return ok ? 0 : 2;
                }
                catch (Exception ex)
                {
                    try
                    {
                        Directory.CreateDirectory(dir);
                        File.WriteAllText(Path.Combine(dir, "selftest.log"), "EXCEPTION: " + ex);
                    }
                    catch
                    {
                    }
                    Environment.Exit(3);
                    return 3;
                }
            }

            Application.ThreadException += delegate(object sender, System.Threading.ThreadExceptionEventArgs e)
            {
                string log = CrashLog.Write("Interfaz", e.Exception);
                MessageBox.Show("Se produjo un error inesperado.\r\n\r\n" + e.Exception.Message
                    + (string.IsNullOrEmpty(log) ? string.Empty : "\r\n\r\nRegistro: " + log),
                    "Mini Word", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
            {
                CrashLog.Write("Aplicación", e.ExceptionObject as Exception);
            };
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            try
            {
                Application.Run(new NativeEditorForm(args));
                return 0;
            }
            catch (Exception ex)
            {
                string log = CrashLog.Write("Inicio", ex);
                MessageBox.Show("No se pudo iniciar Mini Word.\r\n\r\n" + ex.Message
                    + (string.IsNullOrEmpty(log) ? string.Empty : "\r\n\r\nRegistro: " + log),
                    "Mini Word", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }
    }
}
