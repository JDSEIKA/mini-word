using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;

namespace WordPadClone.Formats
{
    /// <summary>
    /// Reader for Word 97-2003 binary documents (MS-DOC).  It reads the OLE2
    /// container directly, reconstructs the piece table, and applies the
    /// character/paragraph formatting stored in the Word binary FKPs.
    /// No Office installation, COM object or third-party library is required.
    /// </summary>
    internal static class DocFormat
    {
        private const int MaxCharacters = 8000000;

        public static DocumentModel Read(string path)
        {
            CompoundFile cfb = CompoundFile.Open(path);
            byte[] word = cfb.ReadStream("WordDocument");
            if (word.Length < 0x200) throw new InvalidDataException("El flujo WordDocument está truncado o no es válido.");

            int ident = ReadU16(word, 0);
            if (ident != 0xA5EC) throw new InvalidDataException("El archivo no contiene un documento de Word válido.");

            int nFib = ReadU16(word, 2);
            if (nFib < 0x00C1 || nFib > 0x0120)
                throw new InvalidDataException("Formato de Word anterior a Word 97 no compatible. Guarde el documento como .docx o .doc de Word 97-2003.");

            int flags = ReadU16(word, 0x0A);
            if ((flags & 0x0100) != 0)
                throw new InvalidDataException("El documento está cifrado o protegido con contraseña. Quite la contraseña en Word y vuelva a intentarlo.");

            bool whichTable = (flags & 0x0200) != 0;
            string tableName = whichTable ? "1Table" : "0Table";
            if (!cfb.HasStream(tableName)) tableName = whichTable ? "0Table" : "1Table";
            if (!cfb.HasStream(tableName)) throw new InvalidDataException("El documento no contiene la tabla de formato (0Table/1Table).");
            byte[] table = cfb.ReadStream(tableName);

            int ccpText = ReadInt(word, 0x4C);
            if (ccpText < 0 || ccpText > MaxCharacters)
                throw new InvalidDataException("El documento contiene demasiado texto para abrirlo en este editor.");

            DocReader reader = new DocReader(cfb, word, table, ccpText);
            return reader.Read();
        }

        // ------------------------------------------------------------------
        // All parsing is kept in one reader so resource lifetime is simple.
        // ------------------------------------------------------------------

        private sealed class DocReader
        {
            private readonly CompoundFile _cfb;
            private readonly byte[] _word;
            private readonly byte[] _table;
            private readonly int _ccpText;
            private byte[] _data;
            private readonly List<Piece> _pieces = new List<Piece>();
            private readonly List<DocChar> _chars = new List<DocChar>();
            private readonly List<ChpxRun> _chpx = new List<ChpxRun>();
            private readonly List<PapxRun> _papx = new List<PapxRun>();
            private readonly List<string> _fonts = new List<string>();
            private readonly List<StyleDef> _styles = new List<StyleDef>();
            private readonly Dictionary<int, ListDef> _lists = new Dictionary<int, ListDef>();
            private readonly Dictionary<int, int[]> _listCounters = new Dictionary<int, int[]>();
            private RunFormat _baseRun = new RunFormat("Times New Roman", 20);
            private SectionInfo _section = new SectionInfo();

            public DocReader(CompoundFile cfb, byte[] word, byte[] table, int ccpText)
            {
                _cfb = cfb;
                _word = word;
                _table = table;
                _ccpText = ccpText;
                try { if (cfb.HasStream("Data")) _data = cfb.ReadStream("Data"); }
                catch { _data = null; }
            }

            private sealed class Piece
            {
                public int CpStart;
                public int CpEnd;
                public int FcStart;
                public bool Compressed;
            }

            private sealed class DocChar
            {
                public char Ch;
                public int Fc;
            }

            private sealed class ChpxRun
            {
                public int FcStart;
                public int FcEnd;
                public byte[] Grpprl;
            }

            private sealed class PapxRun
            {
                public int FcStart;
                public int FcEnd;
                public int Istd;
                public byte[] Grpprl;
            }

            private sealed class StyleDef
            {
                public int Id;
                public int Sti;
                public int Stk;
                public int BaseId;
                public int Cupx;
                public string Name;
                public byte[] Papx;
                public byte[] Chpx;
            }

            private sealed class SectionInfo
            {
                public int PageWidth = 12240;
                public int PageHeight = 15840;
                public int MarginLeft = 1440;
                public int MarginRight = 1440;
                public int MarginTop = 1440;
                public int MarginBottom = 1440;
            }

            private sealed class LevelDef
            {
                public int Start = 1;
                public int Nfc = 0;
                public int Follow = 1;
                public string Xst = string.Empty;
                public byte[] Papx;
                public byte[] Chpx;
            }

            private sealed class ListDef
            {
                public readonly List<LevelDef> Levels = new List<LevelDef>();
            }

            public DocumentModel Read()
            {
                ParsePieces();
                ParseFonts();
                ParseStyles();
                ParseLists();
                ParseFormatting();
                ParseSection();
                DocumentModel model = BuildModel();
                ParseHeaderFooter(model);
                ParseTextboxes(model);
                ParseInlineImages(model);
                model.Title = ReadSummaryProperty(2);
                model.Author = ReadSummaryProperty(4);
                return model;
            }

            // ----------------------------------------------------------
            // Text and piece table
            // ----------------------------------------------------------

            private void ParsePieces()
            {
                int fcClx = ReadInt(_word, 0x01A2);
                int lcbClx = ReadInt(_word, 0x01A6);
                if (fcClx <= 0 || lcbClx <= 0 || fcClx + lcbClx > _table.Length)
                    throw new InvalidDataException("La tabla de piezas de texto del documento no es válida.");

                int pos = fcClx;
                int end = fcClx + lcbClx;
                int plcStart = -1;
                int plcSize = 0;
                while (pos < end)
                {
                    byte type = _table[pos++];
                    if (type == 1)
                    {
                        if (pos + 2 > end) throw new InvalidDataException("El bloque de propiedades del texto está truncado.");
                        int cb = ReadU16(_table, pos);
                        pos += 2 + cb;
                    }
                    else if (type == 2)
                    {
                        if (pos + 4 > end) throw new InvalidDataException("La tabla de piezas de texto está truncada.");
                        plcSize = ReadInt(_table, pos);
                        plcStart = pos + 4;
                        break;
                    }
                    else
                    {
                        throw new InvalidDataException("La tabla de piezas de texto contiene un tipo desconocido.");
                    }
                    if ((pos & 1) != 0) pos++;
                }

                if (plcStart < 0 || plcSize < 4 || plcStart + plcSize > _table.Length)
                    throw new InvalidDataException("No se encontró la tabla de piezas de texto.");
                int count = (plcSize - 4) / 12;
                if (count <= 0 || count > 200000) throw new InvalidDataException("La tabla de piezas de texto no es válida.");

                int pcdStart = plcStart + (count + 1) * 4;
                for (int i = 0; i < count; i++)
                {
                    int cpStart = ReadInt(_table, plcStart + i * 4);
                    int cpEnd = ReadInt(_table, plcStart + (i + 1) * 4);
                    if (cpStart < 0 || cpEnd < cpStart) continue;
                    uint fcRaw = ReadU32(_table, pcdStart + i * 8 + 2);
                    bool compressed = (fcRaw & 0x40000000) != 0;
                    int fc = (int)(fcRaw & 0x3FFFFFFF);
                    if (compressed) fc /= 2;

                    Piece piece = new Piece();
                    piece.CpStart = cpStart;
                    piece.CpEnd = cpEnd;
                    piece.FcStart = fc;
                    piece.Compressed = compressed;
                    _pieces.Add(piece);
                }

                if (_pieces.Count == 0) throw new InvalidDataException("El documento no contiene texto legible.");
                BuildCharacters();
            }

            private void BuildCharacters()
            {
                int target = _ccpText;
                if (target <= 0) target = int.MaxValue;
                for (int i = 0; i < _pieces.Count && _chars.Count < target; i++)
                {
                    Piece p = _pieces[i];
                    int cpStart = p.CpStart;
                    int cpEnd = Math.Min(p.CpEnd, target);
                    if (cpStart >= target) break;
                    if (cpEnd <= cpStart) continue;
                    int count = cpEnd - cpStart;
                    if (p.Compressed)
                    {
                        if (p.FcStart < 0 || p.FcStart + count > _word.Length)
                            throw new InvalidDataException("El texto del documento apunta fuera del flujo WordDocument.");
                        Encoding encoding = Encoding.GetEncoding(1252);
                        for (int j = 0; j < count; j++)
                        {
                            DocChar c = new DocChar();
                            c.Ch = encoding.GetString(_word, p.FcStart + j, 1)[0];
                            c.Fc = p.FcStart + j;
                            _chars.Add(c);
                        }
                    }
                    else
                    {
                        int bytes = count * 2;
                        if (p.FcStart < 0 || p.FcStart + bytes > _word.Length)
                            throw new InvalidDataException("El texto Unicode del documento apunta fuera del flujo WordDocument.");
                        for (int j = 0; j < count; j++)
                        {
                            DocChar c = new DocChar();
                            c.Ch = (char)ReadU16(_word, p.FcStart + j * 2);
                            c.Fc = p.FcStart + j * 2;
                            _chars.Add(c);
                        }
                    }
                }
                if (_chars.Count == 0) throw new InvalidDataException("El documento no contiene texto legible.");
                if (_chars.Count > MaxCharacters) _chars.RemoveRange(MaxCharacters, _chars.Count - MaxCharacters);
            }

            private void ParseHeaderFooter(DocumentModel model)
            {
                int ccpText = ReadInt(_word, 0x004C);
                int ccpFtn = ReadInt(_word, 0x0050);
                int ccpHdd = ReadInt(_word, 0x0054);
                if (ccpText < 0 || ccpFtn < 0 || ccpHdd <= 0) return;
                int start = ccpText + ccpFtn;
                string text = ReadTextRange(start, start + ccpHdd);
                if (string.IsNullOrEmpty(text)) return;
                text = text.Replace('\u0002', '\n').Replace('\u0003', '\n').Replace('\u0001', ' ');
                string[] paragraphs = text.Replace('\r', '\n').Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                StringBuilder header = new StringBuilder();
                for (int i = 0; i < paragraphs.Length; i++)
                {
                    string line = paragraphs[i].Trim();
                    if (line.Length == 0) continue;
                    if (header.Length > 0) header.Append('\n');
                    header.Append(line);
                }
                model.HeaderText = header.ToString();
                model.SyncFirstSection();
                if (model.Sections.Count > 0) model.Sections[0].HeaderText = model.HeaderText;
            }

            private void ParseTextboxes(DocumentModel model)
            {
                int ccpText = ReadInt(_word, 0x004C);
                int ccpFtn = ReadInt(_word, 0x0050);
                int ccpHdd = ReadInt(_word, 0x0054);
                int ccpMcr = ReadInt(_word, 0x0058);
                int ccpAtn = ReadInt(_word, 0x005C);
                int ccpEdn = ReadInt(_word, 0x0060);
                int ccpTxbx = ReadInt(_word, 0x0064);
                int ccpHdrTxbx = ReadInt(_word, 0x0068);
                if (ccpTxbx <= 0 && ccpHdrTxbx <= 0) return;
                int start = ccpText + ccpFtn + ccpHdd + ccpMcr + ccpAtn + ccpEdn;
                string text = ReadTextRange(start, start + Math.Max(0, ccpTxbx));
                if (string.IsNullOrEmpty(text))
                    text = ReadTextRange(start + Math.Max(0, ccpTxbx), start + Math.Max(0, ccpTxbx) + Math.Max(0, ccpHdrTxbx));
                if (string.IsNullOrEmpty(text)) return;
                text = text.Replace('\r', '\n').Replace('\u0002', '\n').Replace('\u0003', '\n').Replace('\u0001', ' ');
                string[] lines = text.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (line.Length == 0) continue;
                    Paragraph p = new Paragraph();
                    p.Inlines.Add(new TextInline("[Cuadro de texto] " + line, _baseRun.Clone()));
                    model.Paragraphs.Add(p);
                }
            }

            private void ParseInlineImages(DocumentModel model)
            {
                if (_data == null || _data.Length < 8) return;
                int index = 0;
                while (index < _data.Length - 8)
                {
                    string extension = null;
                    int start = -1;
                    int length = 0;
                    if (_data[index] == 0x89 && _data[index + 1] == 0x50 && _data[index + 2] == 0x4E && _data[index + 3] == 0x47)
                    {
                        extension = "png";
                        start = index;
                        for (int i = index + 8; i < _data.Length - 12; i++)
                            if (_data[i] == 'I' && _data[i + 1] == 'E' && _data[i + 2] == 'N' && _data[i + 3] == 'D')
                            { length = i + 12 - start; break; }
                    }
                    else if (_data[index] == 0xFF && _data[index + 1] == 0xD8 && _data[index + 2] == 0xFF)
                    {
                        extension = "jpeg";
                        start = index;
                        for (int i = index + 2; i < _data.Length - 1; i++)
                            if (_data[i] == 0xFF && _data[i + 1] == 0xD9) { length = i + 2 - start; break; }
                    }
                    if (extension == null || length <= 8 || start + length > _data.Length)
                    {
                        index++;
                        continue;
                    }
                    byte[] bytes = new byte[length];
                    Buffer.BlockCopy(_data, start, bytes, 0, length);
                    int width = 2000, height = 1500;
                    ReadImageDimensions(bytes, extension, out width, out height);
                    Paragraph p = new Paragraph();
                    p.Inlines.Add(new ImageInline(bytes, extension, width * 15, height * 15, _baseRun.Clone()));
                    model.Paragraphs.Add(p);
                    index += length;
                }
            }

            private static void ReadImageDimensions(byte[] data, string extension, out int width, out int height)
            {
                width = 2000;
                height = 1500;
                try
                {
                    if (extension == "png" && data.Length > 24)
                    {
                        width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
                        height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
                    }
                    else if (extension == "jpeg")
                    {
                        int pos = 2;
                        while (pos + 9 < data.Length)
                        {
                            if (data[pos] != 0xFF) { pos++; continue; }
                            int marker = data[pos + 1];
                            pos += 2;
                            if (marker == 0xD8 || marker == 0xD9) continue;
                            int size = (data[pos] << 8) | data[pos + 1];
                            if (size < 2 || pos + size > data.Length) break;
                            if ((marker >= 0xC0 && marker <= 0xC3) || (marker >= 0xC5 && marker <= 0xC7)
                                || (marker >= 0xC9 && marker <= 0xCB) || (marker >= 0xCD && marker <= 0xCF))
                            {
                                height = (data[pos + 3] << 8) | data[pos + 4];
                                width = (data[pos + 5] << 8) | data[pos + 6];
                                break;
                            }
                            pos += size;
                        }
                    }
                }
                catch { }
                if (width < 1) width = 2000;
                if (height < 1) height = 1500;
            }

            private string ReadTextRange(int cpStart, int cpEnd)
            {
                if (cpEnd <= cpStart) return string.Empty;
                StringBuilder text = new StringBuilder();
                for (int i = 0; i < _pieces.Count; i++)
                {
                    Piece piece = _pieces[i];
                    int start = Math.Max(cpStart, piece.CpStart);
                    int end = Math.Min(cpEnd, piece.CpEnd);
                    if (end <= start) continue;
                    if (piece.Compressed)
                    {
                        int count = end - start;
                        if (piece.FcStart + (start - piece.CpStart) + count > _word.Length) continue;
                        text.Append(Encoding.GetEncoding(1252).GetString(_word,
                            piece.FcStart + (start - piece.CpStart), count));
                    }
                    else
                    {
                        int first = start - piece.CpStart;
                        for (int j = first; j < first + (end - start); j++)
                        {
                            int offset = piece.FcStart + j * 2;
                            if (offset + 1 < _word.Length) text.Append((char)ReadU16(_word, offset));
                        }
                    }
                }
                return text.ToString();
            }

            // ----------------------------------------------------------
            // Font table and styles
            // ----------------------------------------------------------

            private void ParseFonts()
            {
                int fc = ReadInt(_word, 0x0112);
                int lcb = ReadInt(_word, 0x0116);
                if (fc < 0 || lcb <= 0 || fc + lcb > _table.Length) return;
                int pos = fc;
                int end = fc + lcb;
                if (pos + 4 > end) return;
                int count = ReadU16(_table, pos);
                int extra = ReadU16(_table, pos + 2);
                pos += 4;
                for (int i = 0; i < count && pos < end; i++)
                {
                    int size = _table[pos];
                    pos++;
                    if (size < 41 || pos + size > end) break;
                    int nameOffset = pos + 39;
                    int nameEnd = pos + size;
                    StringBuilder sb = new StringBuilder();
                    while (nameOffset + 1 < nameEnd)
                    {
                        char c = (char)ReadU16(_table, nameOffset);
                        nameOffset += 2;
                        if (c == '\0') break;
                        sb.Append(c);
                    }
                    _fonts.Add(sb.ToString());
                    pos += size + extra;
                }
            }

            private void ParseStyles()
            {
                int fc = ReadInt(_word, 0x00A2);
                int lcb = ReadInt(_word, 0x00A6);
                if (fc < 0 || lcb <= 0 || fc + lcb > _table.Length) return;
                int end = fc + lcb;
                if (fc + 2 > end) return;
                int cbStshi = ReadU16(_table, fc);
                int stshi = fc + 2;
                if (cbStshi < 4 || stshi + cbStshi > end) return;
                int cstd = ReadU16(_table, stshi);
                int cbBase = ReadU16(_table, stshi + 2);
                if (cbBase < 10 || cbBase > 64) cbBase = 10;
                int pos = stshi + cbStshi;
                if ((pos & 1) != 0) pos++;

                for (int i = 0; i < cstd && pos + 2 <= end; i++)
                {
                    int cbStd = ReadU16(_table, pos);
                    pos += 2;
                    StyleDef style = new StyleDef();
                    style.Id = i;
                    if (cbStd == 0)
                    {
                        _styles.Add(style);
                        continue;
                    }
                    if (cbStd < 12 || pos + cbStd > end) break;
                    int std = pos;
                    try
                    {
                        int stdf = ReadU16(_table, std);
                        int stkBase = ReadU16(_table, std + 2);
                        int cupxNext = ReadU16(_table, std + 4);
                        style.Sti = stdf & 0x0FFF;
                        style.Stk = stkBase & 0x000F;
                        style.BaseId = (stkBase >> 4) & 0x0FFF;
                        style.Cupx = cupxNext & 0x000F;

                        int namePos = std + cbBase;
                        if (namePos + 2 <= std + cbStd)
                        {
                            int cch = ReadU16(_table, namePos);
                            if (cch >= 0 && namePos + 2 + cch * 2 <= std + cbStd)
                            {
                                string rawName = Encoding.Unicode.GetString(_table, namePos + 2, cch * 2).TrimEnd('\0');
                                int comma = rawName.IndexOf(',');
                                style.Name = comma >= 0 ? rawName.Substring(0, comma) : rawName;
                            }
                            int dataPos = namePos + 2 + Math.Max(0, cch) * 2;
                            int limit = std + cbStd;
                            if (style.Stk == 1)
                            {
                                int start;
                                int cb;
                                if (FindPapxUpx(dataPos, limit, style.Id, out start, out cb))
                                {
                                    style.Papx = Slice(_table, start + 4, cb - 2);
                                    dataPos = start + 2 + cb;
                                    if ((dataPos & 1) != 0) dataPos++;
                                }
                                int cstart, ccb;
                                if (FindChpxUpx(dataPos, limit, out cstart, out ccb))
                                    style.Chpx = Slice(_table, cstart + 2, ccb);
                            }
                            else if (style.Stk == 2)
                            {
                                int cstart, ccb;
                                if (FindChpxUpx(dataPos, limit, out cstart, out ccb))
                                    style.Chpx = Slice(_table, cstart + 2, ccb);
                            }
                        }
                    }
                    catch
                    {
                        // A damaged optional style must not discard the text.
                    }
                    _styles.Add(style);
                    pos = std + cbStd;
                    if ((pos & 1) != 0) pos++;
                }

                // The Normal style is applied per paragraph (style chain 0), so
                // keep this as the unstyled document default.
                _baseRun = new RunFormat("Times New Roman", 20);
            }

            private bool FindPapxUpx(int from, int limit, int styleId, out int start, out int cb)
            {
                start = -1;
                cb = 0;
                for (int pad = 0; pad <= 6; pad += 2)
                {
                    int pos = from + pad;
                    if (pos + 4 > limit) continue;
                    int candidate = ReadU16(_table, pos);
                    if (candidate < 2 || pos + 2 + candidate > limit) continue;
                    if (ReadU16(_table, pos + 2) != styleId) continue;
                    start = pos;
                    cb = candidate;
                    return true;
                }
                return false;
            }

            private bool FindChpxUpx(int from, int limit, out int start, out int cb)
            {
                start = -1;
                cb = 0;
                for (int pad = 0; pad <= 6; pad += 2)
                {
                    int pos = from + pad;
                    if (pos + 2 > limit) continue;
                    int candidate = ReadU16(_table, pos);
                    if (candidate <= 0 || pos + 2 + candidate > limit) continue;
                    // A CHPX begins with a sprm; reject an all-padding block.
                    if (candidate < 2) continue;
                    start = pos;
                    cb = candidate;
                    return true;
                }
                return false;
            }

            private static byte[] Slice(byte[] data, int offset, int length)
            {
                if (length <= 0 || offset < 0 || offset + length > data.Length) return null;
                byte[] result = new byte[length];
                Buffer.BlockCopy(data, offset, result, 0, length);
                return result;
            }

            private void ParseLists()
            {
                int fcLst = ReadInt(_word, 0x02E2);
                int lcbLst = ReadInt(_word, 0x02E6);
                int fcLfo = ReadInt(_word, 0x02EA);
                int lcbLfo = ReadInt(_word, 0x02EE);
                if (fcLst < 0 || lcbLst < 4 || fcLst + lcbLst > _table.Length) return;

                int cLst = ReadU16(_table, fcLst);
                if (cLst <= 0 || cLst > 2000) return;
                int lvlPos = fcLst + 2 + cLst * 28;
                if (lvlPos > _table.Length) return;
                Dictionary<int, ListDef> byLsid = new Dictionary<int, ListDef>();

                for (int i = 0; i < cLst; i++)
                {
                    int lstf = fcLst + 2 + i * 28;
                    if (lstf + 28 > fcLst + lcbLst) break;
                    int lsid = ReadInt(_table, lstf);
                    int flags = _table[lstf + 26];
                    bool simple = (flags & 1) != 0;
                    int levelsCount = simple ? 1 : 9;
                    ListDef def = new ListDef();
                    for (int l = 0; l < levelsCount && lvlPos + 28 <= _table.Length; l++)
                    {
                        int lvlf = lvlPos;
                        LevelDef level = new LevelDef();
                        level.Start = ReadInt(_table, lvlf);
                        level.Nfc = _table[lvlf + 4];
                        level.Follow = _table[lvlf + 15];
                        int cbChpx = _table[lvlf + 24];
                        int cbPapx = _table[lvlf + 25];
                        int data = lvlf + 28;
                        if (data + cbPapx + cbChpx + 2 > _table.Length) break;
                        if (cbPapx > 0)
                        {
                            level.Papx = new byte[cbPapx];
                            Buffer.BlockCopy(_table, data, level.Papx, 0, cbPapx);
                        }
                        data += cbPapx;
                        if (cbChpx > 0)
                        {
                            level.Chpx = new byte[cbChpx];
                            Buffer.BlockCopy(_table, data, level.Chpx, 0, cbChpx);
                        }
                        data += cbChpx;
                        int cch = ReadU16(_table, data);
                        data += 2;
                        if (cch > 0 && cch < 256 && data + cch * 2 <= _table.Length)
                        {
                            level.Xst = Encoding.Unicode.GetString(_table, data, cch * 2);
                        }
                        data += Math.Max(0, cch) * 2;
                        lvlPos = data;
                        if ((lvlPos & 1) != 0) lvlPos++;
                        def.Levels.Add(level);
                    }
                    if (!byLsid.ContainsKey(lsid)) byLsid[lsid] = def;
                }

                if (fcLfo < 0 || lcbLfo < 4 || fcLfo + lcbLfo > _table.Length) return;
                int lfoCount = ReadInt(_table, fcLfo);
                if (lfoCount <= 0 || lfoCount > 20000) return;
                if (4 + lfoCount * 16 > lcbLfo) return;
                for (int i = 0; i < lfoCount; i++)
                {
                    int lfo = fcLfo + 4 + i * 16;
                    int lsid = ReadInt(_table, lfo);
                    ListDef def;
                    if (byLsid.TryGetValue(lsid, out def)) _lists[i + 1] = def;
                }
            }

            private List<StyleDef> GetStyleChain(int id)
            {
                List<StyleDef> chain = new List<StyleDef>();
                int guard = 0;
                while (id >= 0 && id < _styles.Count && guard++ < 64)
                {
                    StyleDef style = _styles[id];
                    if (style == null) break;
                    bool already = false;
                    for (int i = 0; i < chain.Count; i++)
                        if (chain[i].Id == id) { already = true; break; }
                    if (already) break;
                    chain.Insert(0, style);
                    id = style.BaseId;
                }
                return chain;
            }

            // ----------------------------------------------------------
            // Character/paragraph formatting (FKP pages)
            // ----------------------------------------------------------

            private void ParseFormatting()
            {
                ParseChpx();
                ParsePapx();
                _chpx.Sort(delegate(ChpxRun a, ChpxRun b) { return a.FcStart.CompareTo(b.FcStart); });
                _papx.Sort(delegate(PapxRun a, PapxRun b) { return a.FcStart.CompareTo(b.FcStart); });
            }

            private void ParseChpx()
            {
                int fc = ReadInt(_word, 0x00FA);
                int lcb = ReadInt(_word, 0x00FE);
                if (fc <= 0 || lcb < 12 || fc + lcb > _table.Length) return;
                int count = (lcb - 4) / 8;
                if (count <= 0) return;
                int pages = fc + (count + 1) * 4;
                for (int i = 0; i < count; i++)
                {
                    uint pn = ReadU32(_table, pages + i * 4) & 0x003FFFFF;
                    if (pn == 0x003FFFFF) continue;
                    int page = (int)pn * 512;
                    if (page < 0 || page + 512 > _word.Length) continue;
                    int crun = _word[page + 511];
                    if (crun <= 0 || (crun + 1) * 4 > 511) continue;
                    int rgb = page + (crun + 1) * 4;
                    for (int r = 0; r < crun; r++)
                    {
                        int start = ReadInt(_word, page + r * 4);
                        int end = ReadInt(_word, page + (r + 1) * 4);
                        int off = _word[rgb + r];
                        if (off == 0 || start >= end) continue;
                        int p = page + off * 2;
                        if (p < 0 || p >= page + 512) continue;
                        int cb = _word[p];
                        if (cb <= 0 || p + 1 + cb > _word.Length) continue;
                        byte[] grpprl = new byte[cb];
                        Buffer.BlockCopy(_word, p + 1, grpprl, 0, cb);
                        ChpxRun run = new ChpxRun();
                        run.FcStart = start;
                        run.FcEnd = end;
                        run.Grpprl = grpprl;
                        _chpx.Add(run);
                    }
                }
            }

            private void ParsePapx()
            {
                int fc = ReadInt(_word, 0x0102);
                int lcb = ReadInt(_word, 0x0106);
                if (fc <= 0 || lcb < 12 || fc + lcb > _table.Length) return;
                int count = (lcb - 4) / 8;
                if (count <= 0) return;
                int pages = fc + (count + 1) * 4;
                for (int i = 0; i < count; i++)
                {
                    uint pn = ReadU32(_table, pages + i * 4) & 0x003FFFFF;
                    if (pn == 0x003FFFFF) continue;
                    int page = (int)pn * 512;
                    if (page < 0 || page + 512 > _word.Length) continue;
                    int crun = _word[page + 511];
                    if (crun <= 0 || (crun + 1) * 4 > 511) continue;
                    int rgb = page + (crun + 1) * 4;
                    if (rgb + crun * 13 > page + 512) continue;
                    for (int r = 0; r < crun; r++)
                    {
                        int start = ReadInt(_word, page + r * 4);
                        int end = ReadInt(_word, page + (r + 1) * 4);
                        int off = _word[rgb + r * 13];
                        if (off == 0 || start >= end) continue;
                        int p = page + off * 2;
                        if (p < 0 || p >= page + 512) continue;
                        int cb = _word[p];
                        int grpStart;
                        int grpLen;
                        if (cb == 0)
                        {
                            if (p + 2 > _word.Length) continue;
                            int cb2 = _word[p + 1];
                            if (cb2 <= 0) continue;
                            grpStart = p + 4;
                            grpLen = cb2 * 2 - 2;
                        }
                        else
                        {
                            grpStart = p + 3;
                            grpLen = cb * 2 - 3;
                        }
                        if (grpLen < 0 || grpStart + grpLen > _word.Length) continue;
                        PapxRun run = new PapxRun();
                        run.FcStart = start;
                        run.FcEnd = end;
                        run.Istd = ReadU16(_word, grpStart - 2);
                        run.Grpprl = new byte[grpLen];
                        if (grpLen > 0) Buffer.BlockCopy(_word, grpStart, run.Grpprl, 0, grpLen);
                        _papx.Add(run);
                    }
                }
            }

            private ChpxRun FindChpx(int fc)
            {
                int lo = 0, hi = _chpx.Count - 1;
                while (lo <= hi)
                {
                    int mid = (lo + hi) / 2;
                    ChpxRun run = _chpx[mid];
                    if (fc < run.FcStart) hi = mid - 1;
                    else if (fc >= run.FcEnd) lo = mid + 1;
                    else return run;
                }
                return null;
            }

            private PapxRun FindPapx(int fc)
            {
                if (_papx.Count == 0) return null;
                int lo = 0, hi = _papx.Count - 1;
                int candidate = -1;
                while (lo <= hi)
                {
                    int mid = (lo + hi) / 2;
                    PapxRun run = _papx[mid];
                    if (run.FcEnd <= fc) lo = mid + 1;
                    else if (run.FcStart > fc) hi = mid - 1;
                    else return run;
                }
                // Paragraph marks are often addressed by the end FC rather than
                // the start.  Fall back to the nearest preceding run.
                lo = 0;
                hi = _papx.Count - 1;
                while (lo <= hi)
                {
                    int mid = (lo + hi) / 2;
                    if (_papx[mid].FcStart <= fc) { candidate = mid; lo = mid + 1; }
                    else hi = mid - 1;
                }
                return candidate >= 0 ? _papx[candidate] : null;
            }

            // ----------------------------------------------------------
            // Property sprms
            // ----------------------------------------------------------

            private delegate void SprmVisitor(int sprm, int offset, int length);

            private static void WalkSprms(byte[] grp, SprmVisitor visitor)
            {
                if (grp == null || grp.Length < 2) return;
                int pos = 0;
                int guard = 0;
                while (pos + 2 <= grp.Length && guard++ < 100000)
                {
                    int sprm = ReadU16(grp, pos);
                    pos += 2;
                    int spra = (sprm >> 13) & 7;
                    int len;
                    switch (spra)
                    {
                        case 0:
                        case 1: len = 1; break;
                        case 2:
                        case 4:
                        case 5: len = 2; break;
                        case 3: len = 4; break;
                        case 7: len = 3; break;
                        default:
                            len = VariableSprmLength(sprm, grp, pos);
                            break;
                    }
                    if (len < 0 || pos + len > grp.Length) break;
                    visitor(sprm, pos, len);
                    pos += len;
                }
            }

            private static int VariableSprmLength(int sprm, byte[] grp, int pos)
            {
                if (pos >= grp.Length) return 0;
                if (sprm == 0xC60D || sprm == 0xC615)
                {
                    int cTabs = grp[pos];
                    return 1 + cTabs * 4;
                }
                if (sprm == 0xD608 || sprm == 0xD606)
                {
                    if (pos + 2 > grp.Length) return 0;
                    return 2 + ReadU16(grp, pos);
                }
                if (sprm == 0xC63E) return 1 + 84;
                if (sprm == 0xCA57 || sprm == 0xC63F)
                {
                    int cb = grp[pos];
                    return 1 + cb;
                }
                return grp[pos] + 1;
            }

            private void ApplyCharSprms(RunFormat f, byte[] grp)
            {
                WalkSprms(grp, delegate(int sprm, int off, int len)
                {
                    int v = len > 0 ? grp[off] : 0;
                    switch (sprm)
                    {
                        case 0x0835: f.Bold = Toggle(f.Bold, v); break;
                        case 0x0836: f.Italic = Toggle(f.Italic, v); break;
                        case 0x0837: f.Strikeout = Toggle(f.Strikeout, v); break;
                        case 0x2A53: if (v != 0) f.Strikeout = true; break;
                        case 0x2A3E: f.Underline = v != 0; break;
                        case 0x2A42:
                            Color? ico = IndexedColor(v);
                            if (ico.HasValue) f.ForeColor = ico;
                            break;
                        case 0x4A43:
                            if (len >= 2) { int hp = ReadU16(grp, off); if (hp > 0 && hp < 4000) f.SizeHalfPoints = hp; }
                            break;
                        case 0x4845:
                            if (len >= 2)
                            {
                                short pos = (short)ReadU16(grp, off);
                                if (pos > 0) f.Script = 1;
                                else if (pos < 0) f.Script = -1;
                                else f.Script = 0;
                            }
                            break;
                        case 0x2A48:
                            if (v == 1) f.Script = 1;
                            else if (v == 2) f.Script = -1;
                            else f.Script = 0;
                            break;
                        case 0x4A4F:
                        case 0x4A51:
                        case 0x4A61:
                            if (len >= 2)
                            {
                                int ftc = ReadU16(grp, off);
                                string name = FontName(ftc);
                                if (!string.IsNullOrEmpty(name)) f.FontFamily = name;
                            }
                            break;
                        case 0x2A0C:
                            {
                                Color? hl = IndexedColor(v);
                                if (hl.HasValue && v != 0) f.BackColor = hl;
                                else if (v == 0) f.BackColor = null;
                            }
                            break;
                        case 0x4866:
                            if (len >= 2)
                            {
                                int shd = ReadU16(grp, off);
                                int icoBack = (shd >> 5) & 0x1F;
                                int ipat = (shd >> 10) & 0x3F;
                                if (icoBack == 0x1F && ipat == 0x3F) f.BackColor = null;
                                else if (ipat != 0 || icoBack != 0)
                                {
                                    Color? shade = IndexedColor(icoBack);
                                    if (shade.HasValue) f.BackColor = shade;
                                }
                            }
                            break;
                        case 0x6870:
                            if (len >= 4)
                            {
                                int color = ReadInt(grp, off);
                                if ((color & 0x01000000) == 0)
                                    f.ForeColor = Color.FromArgb(color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF);
                            }
                            break;
                    }
                });
            }

            private void ApplyParagraphSprms(Paragraph p, byte[] grp, ref int ilfo, ref int ilvl,
                ref bool inTable, ref bool tableRowEnd)
            {
                int listIndex = ilfo;
                int listLevel = ilvl;
                bool table = inTable;
                bool rowEnd = tableRowEnd;
                WalkSprms(grp, delegate(int sprm, int off, int len)
                {
                    int v = len > 0 ? grp[off] : 0;
                    switch (sprm)
                    {
                        case 0x2403:
                        case 0x2461:
                            p.Alignment = (ParagraphAlignment)Math.Min(3, Math.Max(0, v));
                            break;
                        case 0x840F:
                        case 0x845E:
                            if (len >= 2) p.LeftIndentTwips = (short)ReadU16(grp, off);
                            break;
                        case 0x8411:
                        case 0x8460:
                            if (len >= 2) p.FirstLineIndentTwips = (short)ReadU16(grp, off);
                            break;
                        case 0x840E:
                        case 0x845D:
                            if (len >= 2) p.RightIndentTwips = Math.Max(0, (int)(short)ReadU16(grp, off));
                            break;
                        case 0xA413:
                            if (len >= 2) p.SpaceBeforeTwips = Math.Max(0, (int)(short)ReadU16(grp, off));
                            break;
                        case 0xA414:
                            if (len >= 2) p.SpaceAfterTwips = Math.Max(0, (int)(short)ReadU16(grp, off));
                            break;
                        case 0x6412:
                            if (len >= 4)
                            {
                                short dya = (short)ReadU16(grp, off);
                                int mult = ReadU16(grp, off + 2);
                                if (mult != 0)
                                {
                                    if (dya >= 470) p.LineSpacing = 2;
                                    else if (dya >= 350) p.LineSpacing = 1;
                                    else p.LineSpacing = 0;
                                }
                                else p.LineSpacing = 0;
                            }
                            break;
                        case 0x2407:
                            if (v != 0) p.PageBreakBefore = true;
                            break;
                        case 0x460B:
                            if (len >= 2) listIndex = ReadU16(grp, off);
                            break;
                        case 0x260A:
                            if (v >= 0 && v < 9) listLevel = v;
                            break;
                        case 0x2416:
                            table = v != 0;
                            break;
                        case 0x2417:
                            rowEnd = v != 0;
                            break;
                        case 0xC60D:
                        case 0xC615:
                            ApplyTabStops(p, grp, off, len);
                            break;
                    }
                });
                ilfo = listIndex;
                ilvl = listLevel;
                inTable = table;
                tableRowEnd = rowEnd;
            }

            private static void ApplyTabStops(Paragraph p, byte[] grp, int off, int len)
            {
                if (len < 1) return;
                int count = grp[off];
                if (count <= 0 || off + 1 + count * 4 > off + len) return;
                List<int> tabs = new List<int>();
                if (p.TabStopsTwips != null) tabs.AddRange(p.TabStopsTwips);
                int add = off + 1 + count * 2;
                for (int i = 0; i < count; i++)
                {
                    int tab = ReadU16(grp, add + i * 2);
                    if (tab > 0 && !tabs.Contains(tab)) tabs.Add(tab);
                }
                tabs.Sort();
                p.TabStopsTwips = tabs.ToArray();
            }

            private void ApplySectionSprms(SectionInfo s, byte[] grp)
            {
                WalkSprms(grp, delegate(int sprm, int off, int len)
                {
                    switch (sprm)
                    {
                        case 0xB01F: if (len >= 2) s.PageWidth = ReadU16(grp, off); break;
                        case 0xB020: if (len >= 2) s.PageHeight = ReadU16(grp, off); break;
                        case 0xB021: if (len >= 2) s.MarginLeft = ReadU16(grp, off); break;
                        case 0xB022: if (len >= 2) s.MarginRight = ReadU16(grp, off); break;
                        case 0x9023: if (len >= 2) s.MarginTop = (short)ReadU16(grp, off); break;
                        case 0x9024: if (len >= 2) s.MarginBottom = (short)ReadU16(grp, off); break;
                    }
                });
            }

            private static bool Toggle(bool current, int value)
            {
                if (value == 0x80) return current;
                if (value == 0x81) return !current;
                return value != 0;
            }

            private static Color? IndexedColor(int index)
            {
                switch (index)
                {
                    case 0: return null;
                    case 1: return Color.Black;
                    case 2: return Color.Blue;
                    case 3: return Color.Cyan;
                    case 4: return Color.Green;
                    case 5: return Color.Magenta;
                    case 6: return Color.Red;
                    case 7: return Color.Yellow;
                    case 8: return Color.White;
                    case 9: return Color.FromArgb(0, 0, 128);
                    case 10: return Color.FromArgb(0, 128, 128);
                    case 11: return Color.FromArgb(0, 128, 0);
                    case 12: return Color.FromArgb(128, 0, 128);
                    case 13: return Color.FromArgb(128, 0, 0);
                    case 14: return Color.FromArgb(128, 128, 0);
                    case 15: return Color.FromArgb(128, 128, 128);
                    case 16: return Color.FromArgb(192, 192, 192);
                    default: return null;
                }
            }

            private string FontName(int index)
            {
                if (index >= 0 && index < _fonts.Count) return _fonts[index];
                return null;
            }

            // ----------------------------------------------------------
            // Sections and page geometry
            // ----------------------------------------------------------

            private void ParseSection()
            {
                int fc = ReadInt(_word, 0x00CA);
                int lcb = ReadInt(_word, 0x00CE);
                if (fc <= 0 || lcb < 20 || fc + lcb > _table.Length) return;
                int count = (lcb - 4) / 16;
                if (count <= 0) return;
                int sed = fc + (count + 1) * 4;
                int fcSepx = ReadInt(_table, sed + 2);
                if (fcSepx <= 0 || fcSepx + 2 > _word.Length) return;
                int cb = ReadU16(_word, fcSepx);
                if (cb <= 0 || fcSepx + 2 + cb > _word.Length) return;
                byte[] grp = new byte[cb];
                Buffer.BlockCopy(_word, fcSepx + 2, grp, 0, cb);
                SectionInfo s = new SectionInfo();
                ApplySectionSprms(s, grp);
                if (s.PageWidth > 1000 && s.PageWidth < 50000 && s.PageHeight > 1000 && s.PageHeight < 50000)
                    _section = s;
            }

            // ----------------------------------------------------------
            // Model construction
            // ----------------------------------------------------------

            private sealed class ParaBuild
            {
                public readonly List<DocChar> Chars = new List<DocChar>();
                public int MarkFc;
                public bool HasMark;
                public bool PageBreak;
                public bool InTable;
                public bool TableRowEnd;
                public bool IsCellEnd;
            }

            private DocumentModel BuildModel()
            {
                List<ParaBuild> paragraphs = SplitParagraphs();
                DocumentModel model = new DocumentModel();
                model.PageWidthTwips = _section.PageWidth;
                model.PageHeightTwips = _section.PageHeight;
                model.MarginLeftTwips = _section.MarginLeft;
                model.MarginRightTwips = _section.MarginRight;
                model.MarginTopTwips = _section.MarginTop;
                model.MarginBottomTwips = _section.MarginBottom;

                int i = 0;
                while (i < paragraphs.Count)
                {
                    ParaBuild build = paragraphs[i];
                    if (build.InTable || build.IsCellEnd)
                    {
                        List<string[]> rows = new List<string[]>();
                        List<string> row = new List<string>();
                        StringBuilder cell = new StringBuilder();
                        bool sawCell = false;
                        while (i < paragraphs.Count)
                        {
                            ParaBuild cellBlock = paragraphs[i];
                            if (!cellBlock.InTable && !cellBlock.IsCellEnd) break;
                            string value = BlockPlainText(cellBlock);
                            if (cellBlock.IsCellEnd)
                            {
                                if (cell.Length > 0 && value.Length > 0) cell.Append(' ');
                                cell.Append(value);
                                string cellText = cell.ToString().Trim();
                                cell.Length = 0;
                                if (cellBlock.TableRowEnd && cellText.Length == 0 && row.Count > 0)
                                {
                                    // The last \x07 of a row is a row-end marker,
                                    // not an additional empty cell.
                                    rows.Add(row.ToArray());
                                    row.Clear();
                                }
                                else
                                {
                                    row.Add(cellText);
                                    sawCell = true;
                                    if (cellBlock.TableRowEnd)
                                    {
                                        if (row.Count > 0) rows.Add(row.ToArray());
                                        row.Clear();
                                    }
                                }
                            }
                            else
                            {
                                if (cell.Length > 0) cell.Append('\n');
                                cell.Append(value);
                            }
                            i++;
                        }
                        if (row.Count > 0) rows.Add(row.ToArray());
                        if (sawCell && rows.Count > 0)
                        {
                            Paragraph table = new Paragraph();
                            table.TableRows = rows.ToArray();
                            table.TableWidthTwips = Math.Max(1440, _section.PageWidth - _section.MarginLeft - _section.MarginRight);
                            model.Paragraphs.Add(table);
                        }
                        continue;
                    }

                    model.Paragraphs.Add(BuildParagraph(build));
                    i++;
                }

                if (model.Paragraphs.Count == 0) model.Paragraphs.Add(new Paragraph());
                return model;
            }

            private Paragraph BuildParagraph(ParaBuild build)
            {
                Paragraph p = new Paragraph();
                int ilfo = 0, ilvl = 0;
                bool inTable = false, rowEnd = false;
                int istd = 0;
                PapxRun papx = build.HasMark ? FindPapx(build.MarkFc) : null;
                if (papx != null) istd = papx.Istd;

                List<StyleDef> chain = GetStyleChain(istd);
                for (int s = 0; s < chain.Count; s++)
                    ApplyParagraphSprms(p, chain[s].Papx, ref ilfo, ref ilvl, ref inTable, ref rowEnd);
                if (papx != null)
                {
                    ApplyParagraphSprms(p, papx.Grpprl, ref ilfo, ref ilvl, ref inTable, ref rowEnd);
                    int directIstd = FindStyleId(papx.Grpprl);
                    if (directIstd >= 0 && directIstd != istd && directIstd < _styles.Count)
                    {
                        istd = directIstd;
                        chain = GetStyleChain(istd);
                        for (int s = 0; s < chain.Count; s++)
                            ApplyParagraphSprms(p, chain[s].Papx, ref ilfo, ref ilvl, ref inTable, ref rowEnd);
                    }
                }
                p.PageBreakBefore = build.PageBreak || p.PageBreakBefore;

                RunFormat paraBase = _baseRun.Clone();
                for (int s = 0; s < chain.Count; s++)
                    ApplyCharSprms(paraBase, chain[s].Chpx);

                string listPrefix = null;
                if (ilfo > 0)
                {
                    ListDef list;
                    LevelDef level = null;
                    if (_lists.TryGetValue(ilfo, out list) && list != null && ilvl >= 0 && ilvl < list.Levels.Count)
                        level = list.Levels[ilvl];
                    if (level != null)
                    {
                        int dummyList = 0, dummyLevel = 0;
                        bool dummyTable = false, dummyRow = false;
                        ApplyParagraphSprms(p, level.Papx, ref dummyList, ref dummyLevel, ref dummyTable, ref dummyRow);
                        if (level.Nfc == 0x17)
                        {
                            p.Bullet = true;
                            p.BulletLevel = Math.Min(2, Math.Max(0, ilvl));
                        }
                        else if (level.Nfc != 0xFF)
                        {
                            listPrefix = BuildListPrefix(ilfo, ilvl, level);
                        }
                    }
                    else
                    {
                        p.Bullet = true;
                        p.BulletLevel = Math.Min(2, Math.Max(0, ilvl));
                    }
                    if (p.Bullet)
                    {
                        if (p.LeftIndentTwips == 0) p.LeftIndentTwips = 720 * (p.BulletLevel + 1);
                        if (p.FirstLineIndentTwips == 0) p.FirstLineIndentTwips = -360;
                    }
                }

                BuildInlines(build, p, paraBase);
                if (!string.IsNullOrEmpty(listPrefix))
                    p.Inlines.Insert(0, new TextInline(listPrefix, paraBase.Clone()));
                if (p.Inlines.Count == 0)
                    p.Inlines.Add(new TextInline(string.Empty, paraBase.Clone()));
                return p;
            }

            private string BuildListPrefix(int ilfo, int ilvl, LevelDef level)
            {
                int[] counters;
                if (!_listCounters.TryGetValue(ilfo, out counters) || counters == null)
                {
                    counters = new int[9];
                    _listCounters[ilfo] = counters;
                }
                if (ilvl < 0) ilvl = 0;
                if (ilvl > 8) ilvl = 8;
                if (counters[ilvl] < level.Start - 1) counters[ilvl] = level.Start - 1;
                counters[ilvl]++;
                for (int i = ilvl + 1; i < counters.Length; i++) counters[i] = 0;

                string pattern = level.Xst;
                bool hasPlaceholder = false;
                StringBuilder sb = new StringBuilder();
                if (!string.IsNullOrEmpty(pattern))
                {
                    for (int i = 0; i < pattern.Length; i++)
                    {
                        char c = pattern[i];
                        if (c < 9)
                        {
                            hasPlaceholder = true;
                            int value = c < counters.Length ? counters[c] : 0;
                            if (value <= 0) value = 1;
                            sb.Append(FormatListNumber(value, level.Nfc));
                        }
                        else
                        {
                            sb.Append(c);
                        }
                    }
                }
                if (!hasPlaceholder)
                {
                    sb.Append(FormatListNumber(counters[ilvl] <= 0 ? 1 : counters[ilvl], level.Nfc));
                    sb.Append('.');
                }
                if (level.Follow == 0) sb.Append('\t');
                else if (level.Follow == 1) sb.Append(' ');
                return sb.ToString();
            }

            private static string FormatListNumber(int value, int nfc)
            {
                switch (nfc)
                {
                    case 1: return ToRoman(value).ToUpperInvariant();
                    case 2: return ToRoman(value).ToLowerInvariant();
                    case 3: return ToLetters(value).ToUpperInvariant();
                    case 4: return ToLetters(value).ToLowerInvariant();
                    default: return value.ToString(CultureInfo.InvariantCulture);
                }
            }

            private static string ToRoman(int value)
            {
                if (value <= 0 || value > 3999) return value.ToString(CultureInfo.InvariantCulture);
                int[] values = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
                string[] letters = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < values.Length; i++)
                {
                    while (value >= values[i]) { sb.Append(letters[i]); value -= values[i]; }
                }
                return sb.ToString();
            }

            private static string ToLetters(int value)
            {
                if (value <= 0) return value.ToString(CultureInfo.InvariantCulture);
                StringBuilder sb = new StringBuilder();
                while (value > 0)
                {
                    value--;
                    sb.Insert(0, (char)('A' + (value % 26)));
                    value /= 26;
                }
                return sb.ToString();
            }

            private string BlockPlainText(ParaBuild build)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < build.Chars.Count; i++)
                {
                    char ch = build.Chars[i].Ch;
                    if (ch == '\x0B') { sb.Append('\n'); continue; }
                    if (ch == '\x1E') { sb.Append('-'); continue; }
                    if (ch < 32 && ch != '\t') continue;
                    if (ch == '\u00A0') ch = ' ';
                    sb.Append(ch);
                }
                return sb.ToString();
            }

            private void QueryTableFlags(int fc, out bool inTable, out bool rowEnd)
            {
                inTable = false;
                rowEnd = false;
                PapxRun papx = FindPapx(fc);
                if (papx == null) return;
                Paragraph temp = new Paragraph();
                int ilfo = 0, ilvl = 0;
                List<StyleDef> chain = GetStyleChain(papx.Istd);
                for (int s = 0; s < chain.Count; s++)
                    ApplyParagraphSprms(temp, chain[s].Papx, ref ilfo, ref ilvl, ref inTable, ref rowEnd);
                ApplyParagraphSprms(temp, papx.Grpprl, ref ilfo, ref ilvl, ref inTable, ref rowEnd);
            }

            private void BuildInlines(ParaBuild build, Paragraph p, RunFormat baseFormat)
            {
                StringBuilder text = new StringBuilder();
                RunFormat current = null;
                for (int i = 0; i < build.Chars.Count; i++)
                {
                    DocChar dc = build.Chars[i];
                    char ch = dc.Ch;
                    if (ch == '\r' || ch == '\x07' || ch == '\x0C' || ch == '\x0B' || ch == '\f')
                    {
                        if (ch == '\x07')
                        {
                            AppendInline(p, ref text, ref current, "\t", baseFormat);
                        }
                        else if (ch == '\x0B')
                        {
                            AppendInline(p, ref text, ref current, "\n", baseFormat);
                        }
                        continue;
                    }
                    if (ch < 32 && ch != '\t')
                    {
                        if (ch == 0x1E) ch = '-';
                        else if (ch == 0x1F) continue;
                        else if (ch == 0x02) continue;
                        else if (ch == 0x08 || ch == 0x05 || ch == 0x03 || ch == 0x04) continue;
                        else continue;
                    }
                    if (ch == '\u00A0') ch = ' ';

                    RunFormat fmt = baseFormat.Clone();
                    ChpxRun chpx = FindChpx(dc.Fc);
                    if (chpx != null)
                    {
                        int charStyle = FindStyleId(chpx.Grpprl);
                        if (charStyle >= 0 && charStyle < _styles.Count)
                        {
                            List<StyleDef> cchain = GetStyleChain(charStyle);
                            for (int s = 0; s < cchain.Count; s++)
                                ApplyCharSprms(fmt, cchain[s].Chpx);
                        }
                        ApplyCharSprms(fmt, chpx.Grpprl);
                    }
                    AppendInline(p, ref text, ref current, ch.ToString(), fmt);
                }
                FlushInline(p, ref text, ref current);
            }

            private static void AppendInline(Paragraph p, ref StringBuilder text, ref RunFormat current, string value, RunFormat fmt)
            {
                if (string.IsNullOrEmpty(value)) return;
                if (current != null && current.SameAs(fmt)) text.Append(value);
                else
                {
                    FlushInline(p, ref text, ref current);
                    current = fmt.Clone();
                    text.Append(value);
                }
            }

            private static void FlushInline(Paragraph p, ref StringBuilder text, ref RunFormat current)
            {
                if (text.Length > 0 && current != null)
                    p.Inlines.Add(new TextInline(text.ToString(), current));
                text.Length = 0;
                current = null;
            }

            private List<ParaBuild> SplitParagraphs()
            {
                List<ParaBuild> result = new List<ParaBuild>();
                ParaBuild current = new ParaBuild();
                bool pendingBreak = false;
                bool any = false;

                // Field instructions are not visible; only the cached result is
                // shown.  A stack handles nested fields such as IF/PAGE.
                List<FieldState> fields = new List<FieldState>();
                for (int i = 0; i < _chars.Count; i++)
                {
                    DocChar dc = _chars[i];
                    char ch = dc.Ch;
                    if (ch == '\x13')
                    {
                        FieldState fs = new FieldState();
                        fields.Add(fs);
                        continue;
                    }
                    if (ch == '\x14')
                    {
                        if (fields.Count > 0) fields[fields.Count - 1].ShowResult = true;
                        continue;
                    }
                    if (ch == '\x15')
                    {
                        if (fields.Count > 0) fields.RemoveAt(fields.Count - 1);
                        continue;
                    }
                    if (fields.Count > 0)
                    {
                        bool visible = true;
                        for (int f = 0; f < fields.Count; f++)
                            if (!fields[f].ShowResult) { visible = false; break; }
                        if (!visible) { any = true; continue; }
                    }

                    if (ch == '\x0C' || ch == '\f')
                    {
                        if (current.Chars.Count > 0)
                        {
                            current.PageBreak = pendingBreak;
                            AddParagraph(result, current);
                        }
                        current = new ParaBuild();
                        pendingBreak = true;
                        any = true;
                        continue;
                    }
                    if (ch == '\r')
                    {
                        current.MarkFc = dc.Fc;
                        current.HasMark = true;
                        current.PageBreak = pendingBreak;
                        pendingBreak = false;
                        AddParagraph(result, current);
                        current = new ParaBuild();
                        any = true;
                        continue;
                    }
                    if (ch == '\x07')
                    {
                        current.MarkFc = dc.Fc;
                        current.HasMark = true;
                        current.IsCellEnd = true;
                        bool inTable, rowEnd;
                        QueryTableFlags(dc.Fc, out inTable, out rowEnd);
                        current.InTable = true;
                        current.TableRowEnd = rowEnd;
                        AddParagraph(result, current);
                        current = new ParaBuild();
                        any = true;
                        continue;
                    }
                    current.Chars.Add(dc);
                    any = true;
                }
                if (current.Chars.Count > 0 || !any)
                {
                    current.PageBreak = pendingBreak;
                    AddParagraph(result, current);
                }
                if (result.Count == 0) result.Add(new ParaBuild());
                return result;
            }

            private void AddParagraph(List<ParaBuild> result, ParaBuild p)
            {
                if (p == null) return;
                if (p.HasMark && !p.IsCellEnd)
                {
                    bool inTable, rowEnd;
                    QueryTableFlags(p.MarkFc, out inTable, out rowEnd);
                    p.InTable = inTable;
                    p.TableRowEnd = rowEnd;
                }
                result.Add(p);
            }

            private sealed class FieldState
            {
                public bool ShowResult;
            }

            private static int FindStyleId(byte[] grp)
            {
                int result = -1;
                WalkSprms(grp, delegate(int sprm, int off, int len)
                {
                    if (sprm == 0x4600 && len >= 2) result = ReadU16(grp, off);
                    else if (sprm == 0x4A30 && len >= 2) result = ReadU16(grp, off);
                });
                return result;
            }

            // ----------------------------------------------------------
            // Optional OLE SummaryInformation metadata
            // ----------------------------------------------------------

            private string ReadSummaryProperty(int id)
            {
                try
                {
                    if (!_cfb.HasStream("\u0005SummaryInformation")) return string.Empty;
                    byte[] data = _cfb.ReadStream("\u0005SummaryInformation");
                    if (data.Length < 48 || ReadU16(data, 0) != 0xFFFE) return string.Empty;
                    int sections = ReadInt(data, 24);
                    if (sections <= 0) return string.Empty;
                    int offset = ReadInt(data, 44);
                    if (offset < 0 || offset + 8 > data.Length) return string.Empty;
                    int count = ReadInt(data, offset + 4);
                    if (count < 0 || offset + 8 + count * 8 > data.Length) return string.Empty;
                    for (int i = 0; i < count; i++)
                    {
                        int pid = ReadInt(data, offset + 8 + i * 8);
                        int rel = ReadInt(data, offset + 12 + i * 8);
                        if (pid != id || rel < 0 || offset + rel + 8 > data.Length) continue;
                        int type = ReadU16(data, offset + rel);
                        int value = offset + rel + 4;
                        if (type == 30)
                        {
                            int len = ReadInt(data, value);
                            if (len <= 0 || value + 4 + len > data.Length) return string.Empty;
                            return Encoding.Default.GetString(data, value + 4, len).TrimEnd('\0');
                        }
                        if (type == 31)
                        {
                            int chars = ReadInt(data, value);
                            if (chars <= 0 || value + 4 + chars * 2 > data.Length) return string.Empty;
                            return Encoding.Unicode.GetString(data, value + 4, chars * 2).TrimEnd('\0');
                        }
                    }
                }
                catch
                {
                }
                return string.Empty;
            }
        }

        // ------------------------------------------------------------------
        // Primitive readers (bounds checked by callers where it matters)
        // ------------------------------------------------------------------

        private static int ReadU16(byte[] data, int offset)
        {
            if (offset < 0 || offset + 2 > data.Length) return 0;
            return data[offset] | (data[offset + 1] << 8);
        }

        private static int ReadInt(byte[] data, int offset)
        {
            if (offset < 0 || offset + 4 > data.Length) return 0;
            return data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24);
        }

        private static uint ReadU32(byte[] data, int offset)
        {
            return (uint)ReadInt(data, offset);
        }
    }
}
