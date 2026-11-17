using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace WordPadClone.Formats
{
    /// <summary>
    /// A pragmatic RTF reader covering the subset produced by WordPad,
    /// RichEdit and Word: character formatting, paragraph formatting, bullets,
    /// tabs, tables (flattened) and metafile pictures.
    /// </summary>
    public static class RtfReader
    {
        public static DocumentModel Read(string rtf)
        {
            if (string.IsNullOrEmpty(rtf)) return new DocumentModel();
            return new Parser(rtf).Parse();
        }

        private sealed class ParaState
        {
            public ParagraphAlignment Alignment = ParagraphAlignment.Left;
            public int Left;
            public int Right;
            public int First;
            public int SpaceBefore;
            public int SpaceAfter;
            public int Sl;
            public int SlMult;
            public List<int> Tabs = new List<int>();
            public bool Bullet;
            public int BulletLevel;

            public ParaState Clone()
            {
                ParaState s = new ParaState();
                s.Alignment = Alignment;
                s.Left = Left;
                s.Right = Right;
                s.First = First;
                s.SpaceBefore = SpaceBefore;
                s.SpaceAfter = SpaceAfter;
                s.Sl = Sl;
                s.SlMult = SlMult;
                s.Tabs = new List<int>(Tabs);
                s.Bullet = Bullet;
                s.BulletLevel = BulletLevel;
                return s;
            }

            public void Reset()
            {
                Alignment = ParagraphAlignment.Left;
                Left = 0; Right = 0; First = 0;
                SpaceBefore = 0; SpaceAfter = 0;
                Sl = 0; SlMult = 0;
                Tabs.Clear();
                Bullet = false;
                BulletLevel = 0;
            }
        }

        private sealed class Parser
        {
            private readonly string _s;
            private int _i;
            private readonly int _len;
            private readonly DocumentModel _doc = new DocumentModel();
            private Paragraph _para;
            private RunFormat _fmt = new RunFormat();
            private ParaState _ps = new ParaState();
            private readonly Dictionary<int, string> _fonts = new Dictionary<int, string>();
            private readonly Dictionary<int, Color> _colors = new Dictionary<int, Color>();
            private int _defaultFont;
            private int _ucSkip = 1;
            private Encoding _ansi;
            private bool _pendingBullet;
            private int _pendingBulletLevel;
            private bool _inTable;
            private bool _pendingPageBreak;
            private readonly StringBuilder _cellBuffer = new StringBuilder();
            private readonly List<string> _tableCells = new List<string>();
            private readonly List<string[]> _tableRows = new List<string[]>();
            private int _tableWidth;

            public Parser(string rtf)
            {
                _s = rtf;
                _len = rtf.Length;
                _ansi = Encoding.GetEncoding(1252);
            }

            public DocumentModel Parse()
            {
                int first = _s.IndexOf('{');
                if (first < 0)
                {
                    DocumentModel plain = new DocumentModel();
                    Paragraph p = new Paragraph();
                    p.Inlines.Add(new TextInline(_s, new RunFormat()));
                    plain.Paragraphs.Add(p);
                    return plain;
                }

                _i = first;
                ParseGroup();
                if (_para != null)
                    FinishParagraph();
                FlushTable();
                if (_doc.Paragraphs.Count == 0)
                    _doc.Paragraphs.Add(new Paragraph());
                return _doc;
            }

            private void ParseGroup()
            {
                if (_i >= _len || _s[_i] != '{') return;
                _i++;
                bool star = false;
                if (_i + 1 < _len && _s[_i] == '\\' && _s[_i + 1] == '*')
                {
                    star = true;
                    _i += 2;
                }

                string dest = PeekControlWord();

                if (dest != null && (IsRawDestination(dest) || (star && !IsKnownStarDestination(dest))))
                {
                    string raw = ReadRawBody();
                    HandleRaw(dest, raw);
                    return;
                }

                RunFormat savedFmt = _fmt.Clone();
                ParaState savedPs = _ps.Clone();
                ParseContent();
                _fmt = savedFmt;
                _ps = savedPs;
            }

            private void ParseContent()
            {
                while (_i < _len)
                {
                    char c = _s[_i];
                    if (c == '}')
                    {
                        _i++;
                        return;
                    }
                    if (c == '{')
                    {
                        ParseGroup();
                        continue;
                    }
                    if (c == '\\')
                    {
                        _i++;
                        ParseControl();
                        continue;
                    }
                    if (c == '\r' || c == '\n')
                    {
                        _i++;
                        continue;
                    }
                    _i++;
                    AppendChar(c);
                }
            }

            private void ParseControl()
            {
                if (_i >= _len) return;
                char c = _s[_i];

                if (c == '\'')
                {
                    if (_i + 2 < _len)
                    {
                        int b = Hex(_s[_i + 1]) * 16 + Hex(_s[_i + 2]);
                        _i += 3;
                        if (b >= 0) AppendChar(_ansi.GetChars(new byte[] { (byte)b })[0]);
                    }
                    else _i++;
                    return;
                }

                if (c == '\\' || c == '{' || c == '}')
                {
                    _i++;
                    AppendChar(c);
                    return;
                }
                if (c == '~')
                {
                    _i++;
                    AppendChar('\u00A0');
                    return;
                }
                if (c == '_')
                {
                    _i++;
                    AppendChar('\u2011');
                    return;
                }
                if (c == '-')
                {
                    _i++;
                    return;
                }
                if (c == '*' || c == '\n' || c == '\r')
                {
                    _i++;
                    return;
                }
                if (!char.IsLetter(c))
                {
                    _i++;
                    return;
                }

                int wordStart = _i;
                while (_i < _len && char.IsLetter(_s[_i])) _i++;
                string word = _s.Substring(wordStart, _i - wordStart);
                int num = 0;
                bool hasNum = false;
                bool negative = false;
                if (_i < _len && _s[_i] == '-')
                {
                    negative = true;
                    _i++;
                }
                int numStart = _i;
                while (_i < _len && char.IsDigit(_s[_i])) _i++;
                if (_i > numStart)
                {
                    hasNum = true;
                    long n;
                    if (long.TryParse(_s.Substring(numStart, _i - numStart), out n))
                    {
                        if (negative) n = -n;
                        if (n > int.MaxValue) n = int.MaxValue;
                        if (n < int.MinValue) n = int.MinValue;
                        num = (int)n;
                    }
                }
                if (_i < _len && _s[_i] == ' ') _i++;

                HandleWord(word, num, hasNum);
            }

            private void HandleWord(string word, int num, bool hasNum)
            {
                switch (word)
                {
                    case "fonttbl":
                    case "colortbl":
                    case "pict":
                    case "pn":
                    case "pntext":
                        // These are normally handled as raw groups before we get here.
                        break;

                    case "ansicpg":
                        if (hasNum) SetAnsiCodePage(num);
                        break;
                    case "deff":
                        if (hasNum) _defaultFont = num;
                        break;
                    case "uc":
                        if (hasNum) _ucSkip = Math.Max(0, num);
                        break;
                    case "u":
                        if (hasNum)
                        {
                            int value = num < 0 ? num + 65536 : num;
                            AppendChar((char)value);
                            SkipFallback();
                        }
                        break;

                    case "f":
                        if (hasNum)
                        {
                            string name;
                            if (_fonts.TryGetValue(num, out name)) _fmt.FontFamily = name;
                        }
                        break;
                    case "fs":
                        if (hasNum && num > 0) _fmt.SizeHalfPoints = num;
                        break;
                    case "b":
                        _fmt.Bold = !hasNum || num != 0;
                        break;
                    case "i":
                        _fmt.Italic = !hasNum || num != 0;
                        break;
                    case "ul":
                        _fmt.Underline = !hasNum || num != 0;
                        break;
                    case "ulnone":
                        _fmt.Underline = false;
                        break;
                    case "strike":
                    case "striked":
                        _fmt.Strikeout = !hasNum || num != 0;
                        break;
                    case "cf":
                        if (hasNum)
                        {
                            if (num == 0) _fmt.ForeColor = null;
                            else
                            {
                                Color col;
                                if (_colors.TryGetValue(num, out col)) _fmt.ForeColor = col;
                            }
                        }
                        break;
                    case "highlight":
                        if (hasNum)
                        {
                            if (num == 0) _fmt.BackColor = null;
                            else
                            {
                                Color col;
                                if (_colors.TryGetValue(num, out col)) _fmt.BackColor = col;
                            }
                        }
                        break;
                    case "super":
                        _fmt.Script = 1;
                        break;
                    case "sub":
                        _fmt.Script = -1;
                        break;
                    case "nosupersub":
                        _fmt.Script = 0;
                        break;
                    case "up":
                        if (hasNum && num != 0) _fmt.Script = 1;
                        break;
                    case "dn":
                        if (hasNum && num != 0) _fmt.Script = -1;
                        break;
                    case "plain":
                        _fmt = new RunFormat();
                        if (_fonts.ContainsKey(_defaultFont)) _fmt.FontFamily = _fonts[_defaultFont];
                        break;

                    case "pard":
                        _ps.Reset();
                        _pendingBullet = false;
                        _pendingBulletLevel = 0;
                        break;
                    case "ql":
                        _ps.Alignment = ParagraphAlignment.Left;
                        break;
                    case "qc":
                        _ps.Alignment = ParagraphAlignment.Center;
                        break;
                    case "qr":
                        _ps.Alignment = ParagraphAlignment.Right;
                        break;
                    case "qj":
                        _ps.Alignment = ParagraphAlignment.Justify;
                        break;
                    case "li":
                        if (hasNum) _ps.Left = num;
                        break;
                    case "ri":
                        if (hasNum) _ps.Right = num;
                        break;
                    case "fi":
                        if (hasNum) _ps.First = num;
                        break;
                    case "sb":
                        if (hasNum) _ps.SpaceBefore = num;
                        break;
                    case "sa":
                        if (hasNum) _ps.SpaceAfter = num;
                        break;
                    case "sl":
                        if (hasNum) _ps.Sl = num;
                        break;
                    case "slmult":
                        if (hasNum) _ps.SlMult = num;
                        break;
                    case "tx":
                        if (hasNum) _ps.Tabs.Add(num);
                        break;
                    case "tabs":
                        break;

                    case "trowd":
                        _tableWidth = 0;
                        break;
                    case "intbl":
                        _inTable = true;
                        break;
                    case "itap":
                        if (hasNum && num > 0) _inTable = true;
                        break;
                    case "cellx":
                        if (hasNum && num > _tableWidth) _tableWidth = num;
                        break;
                    case "par":
                        FinishParagraph();
                        break;
                    case "page":
                        _pendingPageBreak = true;
                        break;
                    case "row":
                        EndTableRow();
                        break;
                    case "line":
                        AppendChar('\n');
                        break;
                    case "tab":
                        AppendChar('\t');
                        break;
                    case "cell":
                        EndTableCell();
                        break;
                    case "bullet":
                        AppendChar('\u2022');
                        break;
                    case "emdash":
                        AppendChar('\u2014');
                        break;
                    case "endash":
                        AppendChar('\u2013');
                        break;
                    case "lquote":
                        AppendChar('\u2018');
                        break;
                    case "rquote":
                        AppendChar('\u2019');
                        break;
                    case "ldblquote":
                        AppendChar('\u201C');
                        break;
                    case "rdblquote":
                        AppendChar('\u201D');
                        break;

                    case "paperw":
                        if (hasNum && num > 0) _doc.PageWidthTwips = num;
                        break;
                    case "paperh":
                        if (hasNum && num > 0) _doc.PageHeightTwips = num;
                        break;
                    case "margl":
                        if (hasNum && num >= 0) _doc.MarginLeftTwips = num;
                        break;
                    case "margr":
                        if (hasNum && num >= 0) _doc.MarginRightTwips = num;
                        break;
                    case "margt":
                        if (hasNum && num >= 0) _doc.MarginTopTwips = num;
                        break;
                    case "margb":
                        if (hasNum && num >= 0) _doc.MarginBottomTwips = num;
                        break;
                    case "pgwsxn":
                        if (hasNum && num > 0) _doc.PageWidthTwips = num;
                        break;
                    case "pghsxn":
                        if (hasNum && num > 0) _doc.PageHeightTwips = num;
                        break;
                    case "marglsxn":
                        if (hasNum) _doc.MarginLeftTwips = num;
                        break;
                    case "margrsxn":
                        if (hasNum) _doc.MarginRightTwips = num;
                        break;
                    case "margtsxn":
                        if (hasNum) _doc.MarginTopTwips = num;
                        break;
                    case "margbsxn":
                        if (hasNum) _doc.MarginBottomTwips = num;
                        break;
                    case "bin":
                        if (hasNum && num > 0) _i = Math.Min(_len, _i + num);
                        break;
                }
            }

            private void SkipFallback()
            {
                int n = _ucSkip;
                while (n > 0 && _i < _len)
                {
                    if (_s[_i] == '?')
                    {
                        _i++;
                    }
                    else if (_s[_i] == '\\' && _i + 1 < _len && _s[_i + 1] == '\'')
                    {
                        _i += 4;
                    }
                    else if (_s[_i] == '\\')
                    {
                        _i++;
                        while (_i < _len && char.IsLetter(_s[_i])) _i++;
                        while (_i < _len && char.IsDigit(_s[_i])) _i++;
                        if (_i < _len && _s[_i] == ' ') _i++;
                    }
                    else
                    {
                        _i++;
                    }
                    n--;
                }
            }

            private void AppendChar(char c)
            {
                if (c == '\0') return;
                Paragraph p = EnsureParagraph();
                if (p.Inlines.Count > 0)
                {
                    TextInline last = p.Inlines[p.Inlines.Count - 1] as TextInline;
                    if (last != null && last.Format.SameAs(_fmt))
                    {
                        last.Text += c;
                        return;
                    }
                }
                p.Inlines.Add(new TextInline(c.ToString(), _fmt.Clone()));
            }

            private Paragraph EnsureParagraph()
            {
                if (_para != null) return _para;
                if (!_inTable) FlushTable();
                _para = new Paragraph();
                ApplyParaState(_para);
                if (_pendingPageBreak)
                {
                    _para.PageBreakBefore = true;
                    _pendingPageBreak = false;
                }
                return _para;
            }

            private void EndTableCell()
            {
                if (!_inTable)
                {
                    AppendChar('\t');
                    return;
                }
                if (_para != null)
                {
                    AppendParagraphToCell(_para);
                    _para = null;
                }
                string text = _cellBuffer.ToString().TrimEnd('\r', '\n');
                _cellBuffer.Length = 0;
                _tableCells.Add(text);
                _pendingBullet = false;
                _pendingBulletLevel = 0;
            }

            private void EndTableRow()
            {
                if (_inTable)
                {
                    if (_para != null || _cellBuffer.Length > 0)
                        EndTableCell();
                    if (_tableCells.Count > 0)
                    {
                        _tableRows.Add(_tableCells.ToArray());
                        _tableCells.Clear();
                    }
                    _inTable = false;
                }
                else
                {
                    FinishParagraph();
                }
            }

            private void AppendParagraphToCell(Paragraph p)
            {
                for (int i = 0; i < p.Inlines.Count; i++)
                {
                    TextInline t = p.Inlines[i] as TextInline;
                    if (t != null) _cellBuffer.Append(t.Text);
                }
            }

            private void FlushTable()
            {
                if (_tableRows.Count == 0) return;
                Paragraph table = new Paragraph();
                table.TableRows = _tableRows.ToArray();
                table.TableWidthTwips = _tableWidth > 0 ? _tableWidth : 8640;
                _doc.Paragraphs.Add(table);
                _tableRows.Clear();
                _tableCells.Clear();
                _cellBuffer.Length = 0;
                _tableWidth = 0;
            }

            private void ApplyParaState(Paragraph p)
            {
                p.Alignment = _ps.Alignment;
                p.LeftIndentTwips = _ps.Left;
                p.RightIndentTwips = _ps.Right;
                p.FirstLineIndentTwips = _ps.First;
                p.SpaceBeforeTwips = _ps.SpaceBefore;
                p.SpaceAfterTwips = _ps.SpaceAfter;
                if (_ps.SlMult > 0 && _ps.Sl > 0)
                {
                    double mult = _ps.Sl / 240.0;
                    if (mult >= 1.8) p.LineSpacing = 2;
                    else if (mult >= 1.3) p.LineSpacing = 1;
                }
                if (_ps.Tabs.Count > 0)
                {
                    _ps.Tabs.Sort();
                    List<int> unique = new List<int>();
                    for (int i = 0; i < _ps.Tabs.Count; i++)
                        if (unique.Count == 0 || unique[unique.Count - 1] != _ps.Tabs[i]) unique.Add(_ps.Tabs[i]);
                    p.TabStopsTwips = unique.ToArray();
                }

                if (_pendingBullet)
                {
                    p.Bullet = true;
                    p.BulletLevel = _pendingBulletLevel;
                    if (p.LeftIndentTwips == 0) p.LeftIndentTwips = 720 * (p.BulletLevel + 1);
                    if (p.FirstLineIndentTwips == 0) p.FirstLineIndentTwips = -360;
                }
            }

            private void FinishParagraph()
            {
                if (_inTable)
                {
                    if (_para != null)
                    {
                        AppendParagraphToCell(_para);
                        _cellBuffer.Append('\n');
                        _para = null;
                    }
                    _pendingBullet = false;
                    _pendingBulletLevel = 0;
                    return;
                }
                Paragraph p = EnsureParagraph();
                ApplyParaState(p);
                _doc.Paragraphs.Add(p);
                _para = null;
                _pendingBullet = false;
                _pendingBulletLevel = 0;
            }

            private void HandleRaw(string dest, string raw)
            {
                if (dest == null) return;
                switch (dest)
                {
                    case "fonttbl":
                        ParseFontTable(raw);
                        break;
                    case "colortbl":
                        ParseColorTable(raw);
                        break;
                    case "pict":
                        ParsePicture(raw);
                        break;
                    case "pn":
                        ParseListInfo(raw);
                        break;
                    case "pntext":
                        break;
                    case "fldinst":
                        break;
                    case "nonshppict":
                        break;
                    case "stylesheet":
                    case "info":
                    case "listtable":
                    case "listoverridetable":
                    case "generator":
                    case "revtbl":
                    case "rsidtbl":
                    case "themedata":
                    case "colorschememapping":
                    case "latentstyles":
                    case "datastore":
                    case "xmlnstbl":
                    case "header":
                    case "footer":
                    case "object":
                    case "objdata":
                    case "objclass":
                    case "shpinst":
                    case "sp":
                    case "sn":
                    case "sv":
                    case "mmathPr":
                        break;
                }
            }

            private void ParseFontTable(string raw)
            {
                string[] parts = raw.Split(';');
                for (int part = 0; part < parts.Length; part++)
                {
                    string entry = parts[part];
                    Match fm = Regex.Match(entry, @"\\f(\d+)");
                    if (!fm.Success) continue;
                    int index;
                    if (!int.TryParse(fm.Groups[1].Value, out index)) continue;

                    string name = entry.Substring(fm.Index + fm.Length);
                    name = Regex.Replace(name, @"\\[a-zA-Z]+-?\d*\s?", " ");
                    name = name.Replace("{", " ").Replace("}", " ");
                    name = name.Replace("\\\\", "\\").Replace("\\{", "{").Replace("\\}", "}");
                    name = name.Trim();
                    if (name.Length > 0 && !_fonts.ContainsKey(index)) _fonts[index] = name;
                }
            }

            private void ParseColorTable(string raw)
            {
                string[] parts = raw.Split(';');
                int index = 0;
                for (int i = 0; i < parts.Length; i++)
                {
                    string part = parts[i];
                    if (i == 0 && part.IndexOf("colortbl", StringComparison.Ordinal) >= 0 && part.IndexOf("red", StringComparison.Ordinal) < 0)
                    {
                        index++;
                        continue;
                    }
                    Match r = Regex.Match(part, @"\\red(\d+)");
                    Match g = Regex.Match(part, @"\\green(\d+)");
                    Match b = Regex.Match(part, @"\\blue(\d+)");
                    if (r.Success && g.Success && b.Success)
                    {
                        int rv, gv, bv;
                        int.TryParse(r.Groups[1].Value, out rv);
                        int.TryParse(g.Groups[1].Value, out gv);
                        int.TryParse(b.Groups[1].Value, out bv);
                        _colors[index] = Color.FromArgb(Clamp(rv), Clamp(gv), Clamp(bv));
                    }
                    index++;
                }
            }

            private static int Clamp(int v)
            {
                if (v < 0) return 0;
                if (v > 255) return 255;
                return v;
            }

            private void ParseListInfo(string raw)
            {
                bool bullet = raw.IndexOf("pnlvlblt", StringComparison.OrdinalIgnoreCase) >= 0
                    || raw.IndexOf("pnlvlbody", StringComparison.OrdinalIgnoreCase) >= 0
                    || raw.IndexOf("pndec", StringComparison.OrdinalIgnoreCase) >= 0
                    || raw.IndexOf("pnucrm", StringComparison.OrdinalIgnoreCase) >= 0
                    || raw.IndexOf("pnlcltr", StringComparison.OrdinalIgnoreCase) >= 0;
                if (!bullet) return;
                _pendingBullet = true;
                Match m = Regex.Match(raw, @"\\pnlvl(\d)");
                if (m.Success)
                {
                    int lvl;
                    if (int.TryParse(m.Groups[1].Value, out lvl)) _pendingBulletLevel = Math.Min(8, Math.Max(0, lvl));
                }
                if (_para != null && _para.Inlines.Count == 0) ApplyParaState(_para);
            }

            private void ParsePicture(string raw)
            {
                string ext = "wmf";
                if (raw.IndexOf("pngblip", StringComparison.OrdinalIgnoreCase) >= 0) ext = "png";
                else if (raw.IndexOf("jpegblip", StringComparison.OrdinalIgnoreCase) >= 0) ext = "jpeg";
                else if (raw.IndexOf("emfblip", StringComparison.OrdinalIgnoreCase) >= 0) ext = "emf";
                else if (raw.IndexOf("wmetafile", StringComparison.OrdinalIgnoreCase) >= 0) ext = "wmf";
                else if (raw.IndexOf("dibitmap", StringComparison.OrdinalIgnoreCase) >= 0) ext = "bmp";
                else if (raw.IndexOf("wbitmap", StringComparison.OrdinalIgnoreCase) >= 0) ext = "bmp";

                int wGoal = ReadNum(raw, "picwgoal");
                int hGoal = ReadNum(raw, "pichgoal");
                int wPix = ReadNum(raw, "picw");
                int hPix = ReadNum(raw, "pich");
                int width = wGoal > 0 ? wGoal : (wPix > 0 ? wPix * 15 : 1440);
                int height = hGoal > 0 ? hGoal : (hPix > 0 ? hPix * 15 : 1440);

                string clean = Regex.Replace(raw, @"\\blipuid\s+[0-9a-fA-F]+", " ");
                clean = Regex.Replace(clean, @"\\[a-zA-Z]+-?\d*\s?", " ");
                clean = clean.Replace("{", " ").Replace("}", " ").Replace("\\", " ");
                StringBuilder hex = new StringBuilder(clean.Length);
                for (int i = 0; i < clean.Length; i++)
                {
                    char c = clean[i];
                    if ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))
                        hex.Append(c);
                }
                if (hex.Length < 2) return;
                if ((hex.Length & 1) != 0) hex.Length--;

                byte[] data = new byte[hex.Length / 2];
                for (int i = 0; i < data.Length; i++)
                {
                    int hi = Hex(hex[i * 2]);
                    int lo = Hex(hex[i * 2 + 1]);
                    data[i] = (byte)((hi << 4) | lo);
                }
                if (data.Length < 16) return;

                Paragraph p = EnsureParagraph();
                p.Inlines.Add(new ImageInline(data, ext, width, height, _fmt.Clone()));
            }

            private static int ReadNum(string raw, string name)
            {
                Match m = Regex.Match(raw, "\\\\" + name + @"(\d+)");
                if (!m.Success) return 0;
                int v;
                return int.TryParse(m.Groups[1].Value, out v) ? v : 0;
            }

            private string PeekControlWord()
            {
                if (_i >= _len || _s[_i] != '\\') return null;
                int p = _i + 1;
                if (p >= _len) return null;
                int start = p;
                while (p < _len && char.IsLetter(_s[p])) p++;
                if (p == start) return null;
                return _s.Substring(start, p - start);
            }

            private static bool IsRawDestination(string dest)
            {
                switch (dest)
                {
                    case "fonttbl":
                    case "colortbl":
                    case "pict":
                    case "pn":
                    case "pntext":
                    case "fldinst":
                    case "nonshppict":
                    case "stylesheet":
                    case "info":
                    case "listtable":
                    case "listoverridetable":
                    case "generator":
                    case "revtbl":
                    case "rsidtbl":
                    case "themedata":
                    case "colorschememapping":
                    case "latentstyles":
                    case "datastore":
                    case "xmlnstbl":
                    case "header":
                    case "footer":
                    case "object":
                    case "objdata":
                    case "objclass":
                    case "shpinst":
                    case "sp":
                    case "sn":
                    case "sv":
                    case "mmathPr":
                        return true;
                }
                return false;
            }

            private static bool IsKnownStarDestination(string dest)
            {
                return dest == "pn" || dest == "shppict" || dest == "fldinst";
            }

            private string ReadRawBody()
            {
                int start = _i;
                int depth = 1;
                while (_i < _len)
                {
                    char c = _s[_i];
                    if (c == '\\')
                    {
                        _i += 2;
                        continue;
                    }
                    if (c == '{') depth++;
                    else if (c == '}')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            string body = _s.Substring(start, _i - start);
                            _i++;
                            return body;
                        }
                    }
                    _i++;
                }
                return _s.Substring(start);
            }

            private void SetAnsiCodePage(int codePage)
            {
                try
                {
                    _ansi = Encoding.GetEncoding(codePage);
                }
                catch
                {
                    _ansi = Encoding.GetEncoding(1252);
                }
            }

            private static int Hex(char c)
            {
                if (c >= '0' && c <= '9') return c - '0';
                if (c >= 'a' && c <= 'f') return c - 'a' + 10;
                if (c >= 'A' && c <= 'F') return c - 'A' + 10;
                return 0;
            }
        }
    }
}
