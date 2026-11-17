using System;
using System.IO;

namespace WordPadClone.Util
{
    /// <summary>
    /// Writes a file beside its destination and swaps it only after the writer
    /// finishes successfully.  The previous version is kept as a .bak file.
    /// </summary>
    internal static class AtomicFile
    {
        public static void Save(string path, Action<string> writer)
        {
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Ruta vacía.", "path");
            if (writer == null) throw new ArgumentNullException("writer");
            string full = Path.GetFullPath(path);
            string directory = Path.GetDirectoryName(full);
            if (string.IsNullOrEmpty(directory)) throw new InvalidOperationException("La ruta no tiene carpeta.");
            if (!Directory.Exists(directory)) throw new DirectoryNotFoundException(directory);

            string temporary = Path.Combine(directory,
                "." + Path.GetFileName(full) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            bool completed = false;
            try
            {
                writer(temporary);
                if (!File.Exists(temporary))
                    throw new IOException("El escritor no generó el archivo temporal.");
                FileInfo generated = new FileInfo(temporary);
                if (generated.Length <= 0)
                    throw new IOException("El archivo generado está vacío.");

                string backup = full + ".bak";
                if (File.Exists(full))
                {
                    try
                    {
                        File.Replace(temporary, full, backup, true);
                    }
                    catch (Exception replaceError)
                    {
                        if (replaceError is IOException || replaceError is PlatformNotSupportedException
                            || replaceError is UnauthorizedAccessException)
                        {
                            File.Copy(full, backup, true);
                            File.Copy(temporary, full, true);
                            File.Delete(temporary);
                        }
                        else
                        {
                            throw;
                        }
                    }
                }
                else
                {
                    File.Move(temporary, full);
                }
                completed = true;
            }
            finally
            {
                if (!completed && File.Exists(temporary))
                {
                    try { File.Delete(temporary); }
                    catch { }
                }
            }
        }
    }
}
