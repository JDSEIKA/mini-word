using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace WordPadClone.Formats
{
    public sealed class ZipEntryData
    {
        public string Name;
        public byte[] Data;
        public bool Compress;

        public ZipEntryData(string name, byte[] data, bool compress)
        {
            Name = name;
            Data = data;
            Compress = compress;
        }

        public ZipEntryData(string name, string text, bool compress)
            : this(name, new UTF8Encoding(false).GetBytes(text), compress)
        {
        }
    }

    /// <summary>
    /// Minimal ZIP reader/writer.  It keeps the application dependent only on
    /// .NET 4.0 (System.IO.Compression's ZipArchive appeared in 4.5) and lets the
    /// ODT writer store the mandatory first "mimetype" entry without compression.
    /// </summary>
    public static class ZipArchiveLite
    {
        private const uint LocalSig = 0x04034b50;
        private const uint CentralSig = 0x02014b50;
        private const uint EndSig = 0x06054b50;
        private const int MaxPackageBytes = 64 * 1024 * 1024;
        private const int MaxEntries = 4096;
        private const int MaxNameBytes = 1024;

        private static readonly uint[] CrcTable = BuildCrcTable();

        private static uint[] BuildCrcTable()
        {
            uint[] table = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint c = i;
                for (int k = 0; k < 8; k++)
                    c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
                table[i] = c;
            }
            return table;
        }

        private static uint Crc32(byte[] data)
        {
            uint c = 0xFFFFFFFFu;
            for (int i = 0; i < data.Length; i++)
                c = CrcTable[(c ^ data[i]) & 0xFF] ^ (c >> 8);
            return c ^ 0xFFFFFFFFu;
        }

        public static void Write(string path, IList<ZipEntryData> entries)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                Write(fs, entries);
            }
        }

        public static void Write(Stream output, IList<ZipEntryData> entries)
        {
            BinaryWriter bw = new BinaryWriter(output, Encoding.UTF8);
            List<byte[]> centralRecords = new List<byte[]>();

            DateTime now = DateTime.Now;
            ushort dosTime = (ushort)((now.Hour << 11) | (now.Minute << 5) | (now.Second / 2));
            ushort dosDate = (ushort)(((Math.Max(1980, now.Year) - 1980) << 9) | (now.Month << 5) | now.Day);

            for (int i = 0; i < entries.Count; i++)
            {
                ZipEntryData entry = entries[i];
                byte[] name = Encoding.UTF8.GetBytes(entry.Name);
                byte[] data = entry.Data ?? new byte[0];
                uint crc = Crc32(data);
                byte[] payload;
                ushort method;

                if (entry.Compress && data.Length > 0)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Compress, true))
                        {
                            ds.Write(data, 0, data.Length);
                        }
                        payload = ms.ToArray();
                    }
                    method = 8;
                    // Storing incompressible data is smaller and faster.
                    if (payload.Length >= data.Length)
                    {
                        payload = data;
                        method = 0;
                    }
                }
                else
                {
                    payload = data;
                    method = 0;
                }

                long offset = output.Position;
                if (offset > uint.MaxValue) throw new InvalidOperationException("El archivo ZIP es demasiado grande.");

                bw.Write(LocalSig);
                bw.Write((ushort)20);
                bw.Write((ushort)0x0800);
                bw.Write(method);
                bw.Write(dosTime);
                bw.Write(dosDate);
                bw.Write(crc);
                bw.Write((uint)payload.Length);
                bw.Write((uint)data.Length);
                bw.Write((ushort)name.Length);
                bw.Write((ushort)0);
                bw.Write(name);
                bw.Write(payload);

                using (MemoryStream cms = new MemoryStream())
                {
                    BinaryWriter cb = new BinaryWriter(cms);
                    cb.Write(CentralSig);
                    cb.Write((ushort)20);
                    cb.Write((ushort)20);
                    cb.Write((ushort)0x0800);
                    cb.Write(method);
                    cb.Write(dosTime);
                    cb.Write(dosDate);
                    cb.Write(crc);
                    cb.Write((uint)payload.Length);
                    cb.Write((uint)data.Length);
                    cb.Write((ushort)name.Length);
                    cb.Write((ushort)0);
                    cb.Write((ushort)0);
                    cb.Write((ushort)0);
                    cb.Write((ushort)0);
                    cb.Write((uint)0);
                    cb.Write((uint)offset);
                    cb.Write(name);
                    cb.Flush();
                    centralRecords.Add(cms.ToArray());
                }
            }

            uint centralOffset = (uint)output.Position;
            for (int i = 0; i < centralRecords.Count; i++)
                bw.Write(centralRecords[i]);
            uint centralSize = (uint)output.Position - centralOffset;

            bw.Write(EndSig);
            bw.Write((ushort)0);
            bw.Write((ushort)0);
            bw.Write((ushort)entries.Count);
            bw.Write((ushort)entries.Count);
            bw.Write(centralSize);
            bw.Write(centralOffset);
            bw.Write((ushort)0);
            bw.Flush();
        }

        public static Dictionary<string, byte[]> Read(string path)
        {
            FileInfo info = new FileInfo(path);
            if (info.Length > MaxPackageBytes) throw new InvalidDataException("El archivo supera el limite de 64 MB.");
            byte[] all = File.ReadAllBytes(path);
            return Read(all);
        }

        public static Dictionary<string, byte[]> Read(byte[] all)
        {
            if (all == null || all.Length > MaxPackageBytes) throw new InvalidDataException("El archivo supera el limite de 64 MB.");
            int eocd = FindEnd(all);
            if (eocd < 0) throw new InvalidDataException("El archivo ZIP no es válido.");

            int count = BitConverter.ToUInt16(all, eocd + 10);
            uint centralSize = BitConverter.ToUInt32(all, eocd + 12);
            uint centralOffset = BitConverter.ToUInt32(all, eocd + 16);
            if (count > MaxEntries) throw new InvalidDataException("El ZIP contiene demasiadas entradas.");

            if (count == 0xFFFF || centralOffset == 0xFFFFFFFFu)
                throw new InvalidDataException("Los archivos ZIP64 no son compatibles.");

            if (centralOffset > all.Length || centralOffset + centralSize > all.Length)
                throw new InvalidDataException("El directorio central del ZIP está dañado.");

            Dictionary<string, byte[]> result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            long total = 0;
            int pos = (int)centralOffset;

            for (int i = 0; i < count; i++)
            {
                if (pos + 46 > all.Length || BitConverter.ToUInt32(all, pos) != CentralSig)
                    throw new InvalidDataException("Entrada ZIP dañada.");

                ushort flags = BitConverter.ToUInt16(all, pos + 8);
                ushort method = BitConverter.ToUInt16(all, pos + 10);
                uint crc = BitConverter.ToUInt32(all, pos + 16);
                uint compSize = BitConverter.ToUInt32(all, pos + 20);
                uint uncompSize = BitConverter.ToUInt32(all, pos + 24);
                int nameLen = BitConverter.ToUInt16(all, pos + 28);
                int extraLen = BitConverter.ToUInt16(all, pos + 30);
                int commentLen = BitConverter.ToUInt16(all, pos + 32);
                uint localOffset = BitConverter.ToUInt32(all, pos + 42);

                if (pos + 46 + nameLen + extraLen + commentLen > all.Length)
                    throw new InvalidDataException("Entrada ZIP truncada.");

                string name = Encoding.UTF8.GetString(all, pos + 46, nameLen);
                name = name.Replace('\\', '/');
                while (name.StartsWith("./", StringComparison.Ordinal)) name = name.Substring(2);
                if (nameLen == 0 || nameLen > MaxNameBytes || name.IndexOf('\0') >= 0
                    || name.StartsWith("/", StringComparison.Ordinal)
                    || name.IndexOf("../", StringComparison.Ordinal) >= 0
                    || name.IndexOf("..\\", StringComparison.Ordinal) >= 0)
                    throw new InvalidDataException("Nombre de entrada ZIP no seguro.");
                pos += 46 + nameLen + extraLen + commentLen;

                if (uncompSize > 128 * 1024 * 1024)
                    throw new InvalidDataException("Una entrada del ZIP supera el límite de 128 MB.");
                total += uncompSize;
                if (total > 256 * 1024 * 1024)
                    throw new InvalidDataException("El contenido del ZIP supera el límite de 256 MB.");

                if (localOffset + 30 > all.Length || BitConverter.ToUInt32(all, (int)localOffset) != LocalSig)
                    throw new InvalidDataException("Cabecera local ZIP dañada.");

                int lnameLen = BitConverter.ToUInt16(all, (int)localOffset + 26);
                int lextraLen = BitConverter.ToUInt16(all, (int)localOffset + 28);
                long dataStart = (long)localOffset + 30 + lnameLen + lextraLen;
                if (dataStart + compSize > all.Length)
                    throw new InvalidDataException("Datos ZIP truncados.");

                byte[] data;
                if (method == 0)
                {
                    data = new byte[compSize];
                    Buffer.BlockCopy(all, (int)dataStart, data, 0, (int)compSize);
                }
                else if (method == 8)
                {
                    using (MemoryStream input = new MemoryStream(all, (int)dataStart, (int)compSize, false))
                    using (DeflateStream ds = new DeflateStream(input, CompressionMode.Decompress))
                    using (MemoryStream output = new MemoryStream((int)uncompSize))
                    {
                        byte[] buffer = new byte[81920];
                        int n;
                        while ((n = ds.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            output.Write(buffer, 0, n);
                            if (output.Length > 128 * 1024 * 1024)
                                throw new InvalidDataException("Una entrada descomprimida es demasiado grande.");
                        }
                        data = output.ToArray();
                    }
                }
                else
                {
                    throw new InvalidDataException("Método de compresión ZIP no compatible: " + method);
                }

                if (data.Length != uncompSize)
                    throw new InvalidDataException("El tamaño descomprimido no coincide en '" + name + "'.");
                if (Crc32(data) != crc)
                    throw new InvalidDataException("La suma de comprobación CRC no coincide en '" + name + "'.");
                if (!result.ContainsKey(name)) result.Add(name, data);
            }

            return result;
        }

        private static int FindEnd(byte[] data)
        {
            int start = Math.Max(0, data.Length - 65557);
            for (int i = data.Length - 22; i >= start; i--)
            {
                if (BitConverter.ToUInt32(data, i) == EndSig)
                {
                    int commentLen = BitConverter.ToUInt16(data, i + 20);
                    if (i + 22 + commentLen == data.Length) return i;
                }
            }
            return -1;
        }
    }
}
