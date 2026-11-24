using System;
using System.IO;
using System.Text;

namespace WordPadClone.Util
{
    internal static class CrashLog
    {
        public static string Folder
        {
            get
            {
                string overridden = Environment.GetEnvironmentVariable("MINIWORD_LOG_DIR");
                if (!string.IsNullOrEmpty(overridden))
                {
                    try { return Path.GetFullPath(overridden); }
                    catch { }
                }
                return Path.Combine(Path.GetTempPath(), "MiniWord", "Logs");
            }
        }

        public static string Write(string context, Exception exception)
        {
            try
            {
                Directory.CreateDirectory(Folder);
                string path = Path.Combine(Folder, "MiniWord-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
                try
                {
                    FileInfo logInfo = new FileInfo(path);
                    if (logInfo.Exists && logInfo.Length > 2L * 1024 * 1024)
                    {
                        string previous = path + ".1";
                        if (File.Exists(previous)) File.Delete(previous);
                        File.Move(path, previous);
                    }
                }
                catch { }
                StringBuilder text = new StringBuilder();
                text.AppendLine("[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + (context ?? "Error"));
                if (exception != null) text.AppendLine(exception.ToString());
                text.AppendLine();
                File.AppendAllText(path, text.ToString(), new UTF8Encoding(false));
                return path;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
