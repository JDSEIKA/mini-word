using System;
using System.Runtime.InteropServices;

namespace WordPadClone.Util
{
    internal sealed class ParaFormatInfo
    {
        public uint Mask;
        public int StartIndent;   // first line position, twips
        public int RightIndent;
        public int Offset;        // continuation lines relative to first line
        public short Alignment;
        public int[] Tabs = new int[0];
        public int SpaceBefore;
        public int SpaceAfter;
        public int LineSpacing;
        public byte LineSpacingRule;

        public int ContinuationIndent
        {
            get { return StartIndent + Offset; }
        }
    }

    internal static class NativeMethods
    {
        public const int WM_USER = 0x0400;
        public const int EM_SETUNDOLIMIT = WM_USER + 50;
        public const int EM_EMPTYUNDOBUFFER = 0x00CD;
        public const int EM_GETPARAFORMAT = WM_USER + 61;
        public const int EM_SETPARAFORMAT = WM_USER + 71;
        public const int EM_FORMATRANGE = WM_USER + 57;
        public const int EM_SETTARGETDEVICE = WM_USER + 72;
        public const int EM_SETLANGOPTIONS = WM_USER + 120;
        public const int WM_SETREDRAW = 0x000B;
        public const int EM_GETLANGOPTIONS = WM_USER + 121;
        public const int EM_SETEDITSTYLE = WM_USER + 204;
        public const int EM_GETEDITSTYLE = WM_USER + 205;
        public const int EM_GETFIRSTVISIBLELINE = 0x00CE;
        public const int EM_LINESCROLL = 0x00B6;

        public const uint PFM_STARTINDENT = 0x00000001;
        public const uint PFM_RIGHTINDENT = 0x00000002;
        public const uint PFM_OFFSET = 0x00000004;
        public const uint PFM_ALIGNMENT = 0x00000008;
        public const uint PFM_TABSTOPS = 0x00000010;
        public const uint PFM_LINESPACING = 0x00000100;

        public const short PFA_LEFT = 1;
        public const short PFA_RIGHT = 2;
        public const short PFA_CENTER = 3;
        public const short PFA_JUSTIFY = 4;

        public const int IMF_SPELLCHECKING = 0x0800;
        public const int SES_USECTF = 0x00010000;
        public const int SES_CTFALLOWEMBED = 0x00200000;
        public const int SES_CTFALLOWSMARTTAG = 0x00400000;
        public const int SES_CTFALLOWPROOFING = 0x00800000;
        public const int SES_HIDEGRIDLINES = 0x00020000;


        // This RichEdit build accepts the 188-byte PARAFORMAT2 layout (the
        // control reports it through cbSize on EM_GETPARAFORMAT).
        private const int ParaSize = 188;
        private const int ParaSizeAlt = 240;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public RECT(int left, int top, int right, int bottom)
            {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct FORMATRANGE
        {
            public IntPtr hdc;
            public IntPtr hdcTarget;
            public RECT rc;
            public RECT rcPage;
            public int cpMin;
            public int cpMax;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref FORMATRANGE lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hDC, int index);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        public static double GetWindowDpiScale(IntPtr hWnd)
        {
            try
            {
                uint dpi = GetDpiForWindow(hWnd);
                if (dpi > 0) return dpi / 96.0;
            }
            catch
            {
            }
            return GetSystemDpiScale();
        }

        public static double GetSystemDpiScale()
        {
            try
            {
                IntPtr dc = GetDC(IntPtr.Zero);
                if (dc == IntPtr.Zero) return 1.0;
                int dpi = GetDeviceCaps(dc, 88);
                ReleaseDC(IntPtr.Zero, dc);
                return dpi > 0 ? dpi / 96.0 : 1.0;
            }
            catch
            {
                return 1.0;
            }
        }

        public static IntPtr Send(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam)
        {
            return SendMessage(hWnd, msg, wParam, lParam);
        }

        public static IntPtr Send(IntPtr hWnd, int msg, IntPtr wParam, ref FORMATRANGE lParam)
        {
            return SendMessage(hWnd, msg, wParam, ref lParam);
        }

        public static ParaFormatInfo GetParaFormat(IntPtr handle)
        {
            if (handle == IntPtr.Zero) return new ParaFormatInfo();
            ParaFormatInfo info = ReadPara(handle, ParaSize);
            if (info.Mask == 0 && info.StartIndent == 0 && info.RightIndent == 0 && info.Offset == 0)
            {
                ParaFormatInfo alt = ReadPara(handle, ParaSizeAlt);
                if (alt.Mask != 0) return alt;
            }
            return info;
        }

        private static ParaFormatInfo ReadPara(IntPtr handle, int size)
        {
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try
            {
                for (int i = 0; i < 256; i++) Marshal.WriteByte(buffer, i, 0);
                Marshal.WriteInt32(buffer, 0, size);
                SendMessage(handle, EM_GETPARAFORMAT, IntPtr.Zero, buffer);

                ParaFormatInfo info = new ParaFormatInfo();
                info.Mask = (uint)Marshal.ReadInt32(buffer, 4);
                info.StartIndent = Marshal.ReadInt32(buffer, 12);
                info.RightIndent = Marshal.ReadInt32(buffer, 16);
                info.Offset = Marshal.ReadInt32(buffer, 20);
                info.Alignment = Marshal.ReadInt16(buffer, 24);
                int tabCount = Marshal.ReadInt16(buffer, 26);
                if (tabCount < 0) tabCount = 0;
                if (tabCount > 32) tabCount = 32;
                info.Tabs = new int[tabCount];
                for (int i = 0; i < tabCount; i++)
                    info.Tabs[i] = Marshal.ReadInt32(buffer, 28 + i * 4);
                info.SpaceBefore = Marshal.ReadInt32(buffer, 156);
                info.SpaceAfter = Marshal.ReadInt32(buffer, 160);
                info.LineSpacing = Marshal.ReadInt32(buffer, 164);
                info.LineSpacingRule = Marshal.ReadByte(buffer, 170);
                return info;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static void WritePara(IntPtr handle, uint mask, int start, int right, int offset,
            int lineSpacing, byte lineRule, int[] tabs, int size)
        {
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try
            {
                for (int i = 0; i < 256; i++) Marshal.WriteByte(buffer, i, 0);
                Marshal.WriteInt32(buffer, 0, size);
                Marshal.WriteInt32(buffer, 4, unchecked((int)mask));
                if (start >= 0) Marshal.WriteInt32(buffer, 12, start);
                if (right >= 0) Marshal.WriteInt32(buffer, 16, right);
                if (offset >= -100000) Marshal.WriteInt32(buffer, 20, offset);
                if ((mask & PFM_LINESPACING) != 0)
                {
                    Marshal.WriteInt32(buffer, 164, lineSpacing);
                    Marshal.WriteByte(buffer, 170, lineRule);
                }
                if ((mask & PFM_TABSTOPS) != 0 && tabs != null)
                {
                    int count = Math.Min(32, tabs.Length);
                    Marshal.WriteInt16(buffer, 26, (short)count);
                    for (int i = 0; i < count; i++)
                        Marshal.WriteInt32(buffer, 28 + i * 4, tabs[i]);
                }
                SendMessage(handle, EM_SETPARAFORMAT, IntPtr.Zero, buffer);
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        public static void SetIndents(IntPtr handle, int start, int continuation, int right)
        {
            uint mask = 0;
            if (start >= 0) mask |= PFM_STARTINDENT;
            if (right >= 0) mask |= PFM_RIGHTINDENT;
            if (continuation >= -100000) mask |= PFM_OFFSET;

            int startVal = start >= 0 ? start : 0;
            int offsetVal = continuation;
            if (continuation >= -100000 && start < 0)
            {
                ParaFormatInfo current = GetParaFormat(handle);
                startVal = current.StartIndent;
                offsetVal = continuation - startVal;
            }
            else if (continuation >= -100000 && start >= 0)
            {
                offsetVal = continuation - start;
            }
            else
            {
                offsetVal = -100001;
            }

            WritePara(handle, mask, start >= 0 ? start : -1, right >= 0 ? right : -1,
                offsetVal, 0, 0, null, ParaSize);
        }

        public static void SetParagraphAlignment(IntPtr handle, short alignment)
        {
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try
            {
                for (int i = 0; i < 256; i++) Marshal.WriteByte(buffer, i, 0);
                Marshal.WriteInt32(buffer, 0, ParaSize);
                Marshal.WriteInt32(buffer, 4, unchecked((int)PFM_ALIGNMENT));
                Marshal.WriteInt16(buffer, 24, alignment);
                SendMessage(handle, EM_SETPARAFORMAT, IntPtr.Zero, buffer);
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        public static void SetLineSpacing(IntPtr handle, int rule)
        {
            // rule: 0 single, 1 one-and-a-half, 2 double
            WritePara(handle, PFM_LINESPACING, -1, -1, -100001, 0, (byte)rule, null, ParaSize);
        }

        public static void SetTabs(IntPtr handle, int[] tabs)
        {
            WritePara(handle, PFM_TABSTOPS, -1, -1, -100001, 0, 0, tabs ?? new int[0], ParaSize);
        }

        // ------------------------------------------------------------------
        // Printing
        // ------------------------------------------------------------------

        public static int FormatRange(IntPtr handle, IntPtr hdc, NativeMethods.RECT page, NativeMethods.RECT margin, int cpMin, int cpMax)
        {
            FORMATRANGE fr = new FORMATRANGE();
            fr.hdc = hdc;
            fr.hdcTarget = hdc;
            fr.rc = margin;
            fr.rcPage = page;
            fr.cpMin = cpMin;
            fr.cpMax = cpMax;
            return SendMessage(handle, EM_FORMATRANGE, (IntPtr)1, ref fr).ToInt32();
        }

        /// <summary>Paginates without painting; returns the first character of the next page.</summary>
        public static int MeasureRange(IntPtr handle, IntPtr hdc, NativeMethods.RECT page, NativeMethods.RECT margin, int cpMin, int cpMax)
        {
            FORMATRANGE fr = new FORMATRANGE();
            fr.hdc = hdc;
            fr.hdcTarget = hdc;
            fr.rc = margin;
            fr.rcPage = page;
            fr.cpMin = cpMin;
            fr.cpMax = cpMax;
            return SendMessage(handle, EM_FORMATRANGE, IntPtr.Zero, ref fr).ToInt32();
        }

        public static int FirstVisibleLine(IntPtr handle)
        {
            return SendMessage(handle, EM_GETFIRSTVISIBLELINE, IntPtr.Zero, IntPtr.Zero).ToInt32();
        }

        public static void ScrollLines(IntPtr handle, int lines)
        {
            if (lines == 0) return;
            SendMessage(handle, EM_LINESCROLL, IntPtr.Zero, (IntPtr)lines);
        }

        public static void ClearFormatRange(IntPtr handle)
        {
            SendMessage(handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero);
        }
    }
}
