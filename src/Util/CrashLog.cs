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
                if (!string.IsNullOrEmpty(overridden)) return overridden;
                return Path.Combine(Path.GetTempPath(), "MiniWord", "Logs");
            }
        }

        public static string Write(string context, Exception exception)
        {
            try
            {
                Directory.CreateDirectory(Folder);
                string path = Path.Combine(Folder, "MiniWord-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
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
