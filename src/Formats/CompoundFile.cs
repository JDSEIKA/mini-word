using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WordPadClone.Formats
{
    /// <summary>
    /// Minimal, read-only reader for Microsoft Compound File Binary (OLE2/CFB)
    /// containers.  Word 97-2003 documents are CFB files with WordDocument and
    /// 0Table/1Table streams.  The implementation deliberately follows the
    /// published CFB rules instead of using an external library.
    /// </summary>
    internal sealed class CompoundFile
    {
        private const uint FreeSector = 0xFFFFFFFF;
        private const uint EndOfChain = 0xFFFFFFFE;
        private const uint FatSector = 0xFFFFFFFD;
        private const uint DifSector = 0xFFFFFFFC;

        private readonly byte[] _data;
        private readonly int _sectorSize;
        private readonly int _miniSectorSize;
        private readonly int _miniCutoff;
        private readonly uint[] _fat;
        private readonly uint[] _miniFat;
        private readonly Dictionary<string, Entry> _entries;

        private sealed class Entry
        {
            public string Name;
            public byte Type;
            public uint StartSector;
            public long Size;
            public bool IsRoot;
        }

        private CompoundFile(byte[] data, int sectorSize, int miniSectorSize, int miniCutoff,
            uint[] fat, uint[] miniFat, Dictionary<string, Entry> entries)
        {
            _data = data;
            _sectorSize = sectorSize;
            _miniSectorSize = miniSectorSize;
            _miniCutoff = miniCutoff;
            _fat = fat;
            _miniFat = miniFat;
            _entries = entries;
        }

        public static CompoundFile Open(string path)
        {
            byte[] data = File.ReadAllBytes(path);
            if (data.Length < 512) throw new InvalidDataException("El archivo es demasiado pequeño para ser un documento de Word 97-2003.");
            if (!(data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0
                && data[4] == 0xA1 && data[5] == 0xB1 && data[6] == 0x1A && data[7] == 0xE1))
                throw new InvalidDataException("El archivo no es un contenedor OLE2 (Compound File).");

            int major = ReadU16(data, 26);
            int sectorShift = ReadU16(data, 30);
            int miniShift = ReadU16(data, 32);
            if (sectorShift != 9 && sectorShift != 12) throw new InvalidDataException("Tamaño de sector CFB no válido.");
            if (miniShift != 6) throw new InvalidDataException("Tamaño de mini-sector CFB no válido.");
            int sectorSize = 1 << sectorShift;
            int miniSectorSize = 1 << miniShift;
            if (sectorSize < 512 || sectorSize > 4096) throw new InvalidDataException("Sector CFB no compatible.");

            uint firstDir = ReadU32(data, 48);
            int miniCutoff = (int)ReadU32(data, 56);
            uint firstMiniFat = ReadU32(data, 60);
            uint numMiniFat = ReadU32(data, 64);
            uint firstDifat = ReadU32(data, 68);
            uint numDifat = ReadU32(data, 72);
            uint numFat = ReadU32(data, 44);

            if (miniCutoff <= 0) miniCutoff = 4096;

            List<uint> fatSectors = new List<uint>();
            for (int i = 0; i < 109; i++)
            {
                uint sector = ReadU32(data, 76 + i * 4);
                if (sector != FreeSector && sector != EndOfChain) fatSectors.Add(sector);
            }

            // Follow the DIFAT chain, when present.  The last entry of each
            // DIFAT sector points at the next DIFAT sector.
            uint difat = firstDifat;
            int guard = 0;
            while (difat != EndOfChain && difat != FreeSector && guard++ < 100000)
            {
                long off = SectorOffset(difat, sectorSize);
                if (!InRange(data, off, sectorSize)) break;
                int slots = sectorSize / 4;
                for (int i = 0; i < slots - 1; i++)
                {
                    uint sector = ReadU32(data, (int)off + i * 4);
                    if (sector != FreeSector && sector != EndOfChain) fatSectors.Add(sector);
                }
                difat = ReadU32(data, (int)off + (slots - 1) * 4);
            }

            List<uint> fatValues = new List<uint>();
            int fatLimit = Math.Max(1, data.Length / sectorSize + 1);
            if (fatSectors.Count > fatLimit) throw new InvalidDataException("Tabla FAT no válida.");
            for (int i = 0; i < fatSectors.Count; i++)
            {
                long off = SectorOffset(fatSectors[i], sectorSize);
                if (!InRange(data, off, sectorSize)) throw new InvalidDataException("Sector FAT fuera del archivo.");
                for (int j = 0; j < sectorSize; j += 4)
                    fatValues.Add(ReadU32(data, (int)off + j));
            }
            if (numFat > 0 && fatValues.Count == 0) throw new InvalidDataException("No se pudo leer la tabla FAT.");
            uint[] fat = fatValues.ToArray();

            uint[] miniFat = new uint[0];
            if (numMiniFat > 0 && firstMiniFat != EndOfChain && firstMiniFat != FreeSector)
            {
                List<uint> mini = new List<uint>();
                uint sector = firstMiniFat;
                int seen = 0;
                while (sector != EndOfChain && sector != FreeSector && seen++ < fatLimit)
                {
                    long off = SectorOffset(sector, sectorSize);
                    if (!InRange(data, off, sectorSize)) break;
                    for (int j = 0; j < sectorSize; j += 4)
                        mini.Add(ReadU32(data, (int)off + j));
                    sector = NextSector(fat, sector);
                }
                miniFat = mini.ToArray();
            }

            // The directory is a normal FAT stream.  Read every entry first,
            // then collect the children of the root storage.
            byte[] dir = ReadNormalStream(data, fat, firstDir, sectorSize, -1);
            Dictionary<string, Entry> entries = new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);
            List<Entry> all = new List<Entry>();
            for (int off = 0; off + 128 <= dir.Length; off += 128)
            {
                int nameLength = ReadU16(dir, off + 64);
                byte type = dir[off + 66];
                if (type == 0) continue;
                string name = DecodeName(dir, off, nameLength);
                if (string.IsNullOrEmpty(name)) continue;
                Entry e = new Entry();
                e.Name = name;
                e.Type = type;
                e.StartSector = ReadU32(dir, off + 116);
                uint low = ReadU32(dir, off + 120);
                uint high = ReadU32(dir, off + 124);
                e.Size = major >= 4 ? ((long)high << 32) | low : low;
                e.IsRoot = type == 5;
                all.Add(e);
            }

            Entry root = null;
            for (int i = 0; i < all.Count; i++)
                if (all[i].IsRoot) { root = all[i]; break; }
            if (root == null) throw new InvalidDataException("El contenedor OLE2 no tiene un almacén raíz.");

            // Root children are generally the only streams we need.  A few
            // producers use nested storages, so keep the full name too.
            entries[root.Name] = root;
            for (int i = 0; i < all.Count; i++)
            {
                if (all[i].IsRoot) continue;
                if (!entries.ContainsKey(all[i].Name)) entries[all[i].Name] = all[i];
            }
            return new CompoundFile(data, sectorSize, miniSectorSize, miniCutoff, fat, miniFat, entries);
        }

        public bool HasStream(string name)
        {
            return _entries.ContainsKey(name);
        }

        public byte[] ReadStream(string name)
        {
            Entry entry;
            if (!_entries.TryGetValue(name, out entry) || entry.Type != 2)
                throw new InvalidDataException("No se encontró el flujo '" + name + "' en el documento.");
            if (entry.Size < 0 || entry.Size > int.MaxValue) throw new InvalidDataException("Flujo OLE2 demasiado grande.");
            int size = (int)entry.Size;

            if (size < _miniCutoff)
            {
                Entry root;
                if (!_entries.TryGetValue("Root Entry", out root)) throw new InvalidDataException("Falta el almacén raíz OLE2.");
                byte[] miniStream = ReadNormalStream(_data, _fat, root.StartSector, _sectorSize, (int)root.Size);
                if (entry.StartSector == EndOfChain || entry.StartSector == FreeSector)
                    return new byte[0];
                byte[] result = new byte[size];
                int written = 0;
                uint sector = entry.StartSector;
                int guard = 0;
                while (written < size && guard++ < _miniFat.Length + 1 && sector != EndOfChain && sector != FreeSector)
                {
                    long off = (long)sector * _miniSectorSize;
                    if (off < 0 || off >= miniStream.Length) break;
                    int count = Math.Min(_miniSectorSize, size - written);
                    if (off + count > miniStream.Length) count = (int)(miniStream.Length - off);
                    if (count <= 0) break;
                    Buffer.BlockCopy(miniStream, (int)off, result, written, count);
                    written += count;
                    sector = NextSector(_miniFat, sector);
                }
                if (written < size) Array.Resize(ref result, written);
                return result;
            }

            return ReadNormalStream(_data, _fat, entry.StartSector, _sectorSize, size);
        }

        private static byte[] ReadNormalStream(byte[] data, uint[] fat, uint start, int sectorSize, int requested)
        {
            if (start == EndOfChain || start == FreeSector) return new byte[0];
            List<byte[]> blocks = new List<byte[]>();
            int total = 0;
            uint sector = start;
            int guard = 0;
            int maxSectors = fat.Length + 1;
            while (sector != EndOfChain && sector != FreeSector && guard++ < maxSectors)
            {
                long off = SectorOffset(sector, sectorSize);
                if (!InRange(data, off, sectorSize)) break;
                byte[] block = new byte[sectorSize];
                Buffer.BlockCopy(data, (int)off, block, 0, sectorSize);
                blocks.Add(block);
                total += sectorSize;
                if (requested >= 0 && total >= requested) break;
                sector = NextSector(fat, sector);
            }
            int size = requested >= 0 ? Math.Min(requested, total) : total;
            byte[] result = new byte[size];
            int written = 0;
            for (int i = 0; i < blocks.Count && written < size; i++)
            {
                int count = Math.Min(blocks[i].Length, size - written);
                Buffer.BlockCopy(blocks[i], 0, result, written, count);
                written += count;
            }
            return result;
        }

        private static uint NextSector(uint[] fat, uint sector)
        {
            if (sector >= fat.Length) return EndOfChain;
            return fat[sector];
        }

        private static long SectorOffset(uint sector, int sectorSize)
        {
            return ((long)sector + 1) * sectorSize;
        }

        private static bool InRange(byte[] data, long offset, int length)
        {
            return offset >= 0 && length >= 0 && offset + length <= data.Length;
        }

        private static string DecodeName(byte[] data, int offset, int nameLength)
        {
            if (nameLength < 2 || nameLength > 64) return string.Empty;
            int chars = nameLength / 2 - 1;
            if (chars <= 0) return string.Empty;
            return Encoding.Unicode.GetString(data, offset, chars * 2).TrimEnd('\0');
        }

        private static int ReadU16(byte[] data, int offset)
        {
            if (offset < 0 || offset + 2 > data.Length) return 0;
            return data[offset] | (data[offset + 1] << 8);
        }

        private static uint ReadU32(byte[] data, int offset)
        {
            if (offset < 0 || offset + 4 > data.Length) return 0;
            return (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24));
        }
    }
}
