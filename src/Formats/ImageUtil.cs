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
