using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace WordPadClone.Formats
{
    /// <summary>Picture conversion helpers (RichEdit only accepts metafiles).</summary>
    public static class ImageUtil
    {
        public static Image LoadImage(byte[] data, string extension, int widthTwips, int heightTwips)
        {
            if (data == null || data.Length == 0) throw new InvalidDataException("Imagen vacía.");
            if (data.Length > 32 * 1024 * 1024) throw new InvalidDataException("La imagen supera el límite de 32 MB.");
            CheckEncodedDimensions(data, extension);
            if (string.Equals(extension, "wmf", StringComparison.OrdinalIgnoreCase) && !HasPlaceableHeader(data))
            {
                int w = Math.Max(1, widthTwips / 15);
                int h = Math.Max(1, heightTwips / 15);
                byte[] combined = new byte[data.Length + 22];
                BuildPlaceableHeader(w, h).CopyTo(combined, 0);
                data.CopyTo(combined, 22);
                data = combined;
            }
            MemoryStream ms = new MemoryStream(data, false);
            try
            {
                Image img = Image.FromStream(ms);
                if ((long)img.Width * img.Height > 50000000L)
                {
                    img.Dispose();
                    throw new InvalidDataException("La imagen supera el límite de 50 megapíxeles.");
                }
                return img;
            }
            catch
            {
                ms.Dispose();
                throw;
            }
        }

        public static Bitmap Render(ImageInline inline, int maxWidthTwips)
        {
            int wTwips = Math.Max(15, inline.WidthTwips);
            int hTwips = Math.Max(15, inline.HeightTwips);
            if (maxWidthTwips > 0 && wTwips > maxWidthTwips)
            {
                hTwips = (int)Math.Max(15, (long)hTwips * maxWidthTwips / wTwips);
                wTwips = maxWidthTwips;
            }
            int w = Math.Max(1, Math.Min(10000, wTwips / 15));
            int h = Math.Max(1, Math.Min(10000, hTwips / 15));

            using (Image source = LoadImage(inline.Data, inline.Extension, inline.WidthTwips, inline.HeightTwips))
            {
                Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.DrawImage(source, new Rectangle(0, 0, w, h));
                }
                return bmp;
            }
        }

        public static byte[] RenderToPng(ImageInline inline)
        {
            using (Bitmap bmp = Render(inline, 0))
            using (MemoryStream ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }

        /// <summary>Converts metafiles to PNG; leaves portable raster formats untouched.</summary>
        public static ImageInline NormalizeForPackage(ImageInline inline)
        {
            string ext = (inline.Extension ?? "png").ToLowerInvariant();
            if (ext == "wmf" || ext == "emf" || ext == "unknown")
            {
                byte[] png;
                try
                {
                    png = RenderToPng(inline);
                }
                catch
                {
                    return inline;
                }
                return new ImageInline(png, "png", inline.WidthTwips, inline.HeightTwips, inline.Format);
            }
            if (ext == "jpg") ext = "jpeg";
            return new ImageInline(inline.Data, ext, inline.WidthTwips, inline.HeightTwips, inline.Format);
        }

        private static void CheckEncodedDimensions(byte[] data, string extension)
        {
            int width = 0;
            int height = 0;
            string ext = (extension ?? string.Empty).ToLowerInvariant();
            if (data.Length >= 24 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
            {
                width = ReadBigEndianInt(data, 16);
                height = ReadBigEndianInt(data, 20);
            }
            else if (data.Length >= 10 && data[0] == 'G' && data[1] == 'I' && data[2] == 'F')
            {
                width = data[6] | (data[7] << 8);
                height = data[8] | (data[9] << 8);
            }
            else if (data.Length >= 26 && data[0] == 'B' && data[1] == 'M')
            {
                width = Math.Abs(BitConverter.ToInt32(data, 18));
                height = Math.Abs(BitConverter.ToInt32(data, 22));
            }
            else if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xD8)
            {
                int pos = 2;
                while (pos + 9 < data.Length)
                {
                    if (data[pos] != 0xFF) { pos++; continue; }
                    while (pos < data.Length && data[pos] == 0xFF) pos++;
                    if (pos >= data.Length) break;
                    int marker = data[pos++];
                    if (marker == 0xD8 || marker == 0xD9 || marker == 0x01 || marker == 0xDA) continue;
                    if (pos + 2 > data.Length) break;
                    int length = (data[pos] << 8) | data[pos + 1];
                    if (length < 2 || pos + length > data.Length) break;
                    if ((marker >= 0xC0 && marker <= 0xC3)
                        || (marker >= 0xC5 && marker <= 0xC7)
                        || (marker >= 0xC9 && marker <= 0xCB)
                        || (marker >= 0xCD && marker <= 0xCF))
                    {
                        if (pos + 7 < data.Length)
                        {
                            height = (data[pos + 3] << 8) | data[pos + 4];
                            width = (data[pos + 5] << 8) | data[pos + 6];
                        }
                        break;
                    }
                    pos += length;
                }
            }
            if (width > 0 && height > 0 && (long)width * height > 50000000L)
                throw new InvalidDataException("La imagen codificada supera el límite de 50 megapíxeles.");
            if (width > 20000 || height > 20000)
                throw new InvalidDataException("La imagen tiene dimensiones demasiado grandes.");
        }

        private static int ReadBigEndianInt(byte[] data, int offset)
        {
            return (data[offset] << 24) | (data[offset + 1] << 16)
                | (data[offset + 2] << 8) | data[offset + 3];
        }

        private static bool HasPlaceableHeader(byte[] data)
        {
            return data.Length >= 4 && data[0] == 0xD7 && data[1] == 0xCD && data[2] == 0xC6 && data[3] == 0x9A;
        }

        private static byte[] BuildPlaceableHeader(int widthPx, int heightPx)
        {
            int wm = Math.Max(1, widthPx * 1440 / 96);
            int hm = Math.Max(1, heightPx * 1440 / 96);
            byte[] b = new byte[22];
            BitConverter.GetBytes((uint)0x9AC6CDD7).CopyTo(b, 0);
            BitConverter.GetBytes((ushort)0).CopyTo(b, 4);
            BitConverter.GetBytes((short)0).CopyTo(b, 6);
            BitConverter.GetBytes((short)0).CopyTo(b, 8);
            BitConverter.GetBytes((short)Math.Min(short.MaxValue, wm)).CopyTo(b, 10);
            BitConverter.GetBytes((short)Math.Min(short.MaxValue, hm)).CopyTo(b, 12);
            BitConverter.GetBytes((ushort)1440).CopyTo(b, 14);
            BitConverter.GetBytes((uint)0).CopyTo(b, 16);
            ushort sum = 0;
            for (int i = 0; i < 10; i++) sum ^= BitConverter.ToUInt16(b, i * 2);
            BitConverter.GetBytes(sum).CopyTo(b, 20);
            return b;
        }
    }
}
